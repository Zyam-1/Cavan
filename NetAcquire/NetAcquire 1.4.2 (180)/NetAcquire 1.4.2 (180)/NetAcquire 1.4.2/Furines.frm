VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmEditUrines 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Urines"
   ClientHeight    =   8910
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   14265
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form4"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8910
   ScaleWidth      =   14265
   Tag             =   "Urine"
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   12180
      Top             =   -120
   End
   Begin VB.Frame fraSampleID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   150
      TabIndex        =   107
      Top             =   120
      Width           =   2385
      Begin VB.ComboBox cMRU 
         Height          =   315
         Left            =   540
         TabIndex        =   109
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
         TabIndex        =   108
         Top             =   510
         Width           =   1755
      End
      Begin ComCtl2.UpDown udSampleID 
         Height          =   480
         Left            =   1905
         TabIndex        =   110
         Top             =   510
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   847
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "txtSampleID"
         BuddyDispid     =   196612
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
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Urine"
         Height          =   285
         Left            =   660
         TabIndex        =   119
         Top             =   210
         Width           =   825
      End
      Begin VB.Image imgLast 
         Height          =   300
         Left            =   2040
         Picture         =   "Furines.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "Find Last Record"
         Top             =   150
         Width           =   300
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "MRU"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   112
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image iRelevant 
         Height          =   480
         Index           =   1
         Left            =   1530
         Picture         =   "Furines.frx":0442
         Top             =   120
         Width           =   480
      End
      Begin VB.Image iRelevant 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   150
         Picture         =   "Furines.frx":074C
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Left            =   720
         TabIndex        =   111
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   12900
      Picture         =   "Furines.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   8130
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   615
      Left            =   12930
      Picture         =   "Furines.frx":10C0
      Style           =   1  'Graphical
      TabIndex        =   104
      Tag             =   "bprint"
      Top             =   2820
      Width           =   1275
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "&History"
      Height          =   675
      Left            =   12900
      Picture         =   "Furines.frx":172A
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   7410
      Width           =   1275
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   615
      Left            =   12930
      Picture         =   "Furines.frx":1B6C
      Style           =   1  'Graphical
      TabIndex        =   102
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   2160
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Details"
      Enabled         =   0   'False
      Height          =   705
      Left            =   12900
      Picture         =   "Furines.frx":1E76
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   5850
      Width           =   1275
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "&Validate"
      Height          =   765
      Left            =   12900
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Furines.frx":24E0
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   6600
      Width           =   1275
   End
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   765
      Left            =   12900
      Picture         =   "Furines.frx":2922
      Style           =   1  'Graphical
      TabIndex        =   99
      ToolTipText     =   "Log as Phoned"
      Top             =   210
      Width           =   1275
   End
   Begin VB.Frame fra 
      Height          =   5055
      Index           =   3
      Left            =   150
      TabIndex        =   62
      Top             =   3690
      Width           =   2775
      Begin VB.CommandButton cmdNAD 
         Caption         =   "NAD"
         Height          =   285
         Left            =   210
         TabIndex        =   77
         Top             =   180
         Width           =   585
      End
      Begin VB.TextBox txtUrobilinogen 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   76
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox txtBilirubin 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   75
         Top             =   1740
         Width           =   1200
      End
      Begin VB.TextBox txtKetones 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   74
         Top             =   1140
         Width           =   1200
      End
      Begin VB.TextBox txtGlucose 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   73
         Top             =   840
         Width           =   1200
      End
      Begin VB.TextBox txtProtein 
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
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   72
         Top             =   540
         Width           =   1200
      End
      Begin VB.TextBox txtpH 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   71
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtBloodHb 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   70
         Top             =   2040
         Width           =   1200
      End
      Begin VB.TextBox txtRCC 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   69
         Top             =   2640
         Width           =   1200
      End
      Begin VB.TextBox txtWCC 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   68
         Top             =   2340
         Width           =   1200
      End
      Begin VB.ComboBox cmbCasts 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   660
         TabIndex        =   67
         Top             =   3360
         Width           =   2025
      End
      Begin VB.ComboBox cmbCrystals 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   660
         TabIndex        =   66
         Top             =   3030
         Width           =   2025
      End
      Begin VB.ComboBox cmbMisc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   660
         TabIndex        =   65
         Top             =   3690
         Width           =   2025
      End
      Begin VB.ComboBox cmbMisc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   660
         TabIndex        =   64
         Top             =   4020
         Width           =   2025
      End
      Begin VB.ComboBox cmbMisc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   660
         TabIndex        =   63
         Top             =   4350
         Width           =   2025
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Urobilinogen"
         Height          =   195
         Left            =   480
         TabIndex        =   89
         Top             =   1470
         Width           =   885
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Bilirubin"
         Height          =   195
         Left            =   810
         TabIndex        =   88
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Crystals"
         Height          =   195
         Left            =   90
         TabIndex        =   87
         Top             =   3090
         Width           =   540
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Casts"
         Height          =   195
         Left            =   240
         TabIndex        =   86
         Top             =   3420
         Width           =   390
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Ketones"
         Height          =   195
         Left            =   765
         TabIndex        =   85
         Top             =   1170
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Glucose"
         Height          =   195
         Left            =   765
         TabIndex        =   84
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Protein"
         Height          =   195
         Left            =   855
         TabIndex        =   83
         Top             =   570
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "pH"
         Height          =   195
         Index           =   0
         Left            =   1140
         TabIndex        =   82
         Top             =   300
         Width           =   210
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Misc"
         Height          =   195
         Left            =   300
         TabIndex        =   81
         Top             =   3750
         Width           =   330
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Blood/Hb"
         Height          =   195
         Left            =   660
         TabIndex        =   80
         Top             =   2070
         Width           =   690
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "RCC"
         Height          =   195
         Left            =   1020
         TabIndex        =   79
         Top             =   2670
         Width           =   330
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "WCC"
         Height          =   195
         Left            =   960
         TabIndex        =   78
         Top             =   2370
         Width           =   375
      End
   End
   Begin VB.Frame fra 
      Height          =   2385
      Index           =   4
      Left            =   2940
      TabIndex        =   53
      Top             =   3690
      Width           =   2685
      Begin VB.TextBox txtPregnancy 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   57
         ToolTipText     =   "P-Positive N-Negative I-Inconclusive U-Unsuitable"
         Top             =   270
         Width           =   1185
      End
      Begin VB.TextBox txtFatGlobules 
         Height          =   285
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   56
         Top             =   1380
         Width           =   1185
      End
      Begin VB.TextBox txtHCGLevel 
         Height          =   285
         Left            =   1380
         MaxLength       =   5
         TabIndex        =   55
         Top             =   840
         Width           =   1185
      End
      Begin VB.TextBox txtSG 
         Height          =   285
         Left            =   1380
         MaxLength       =   5
         TabIndex        =   54
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Pregnancy Test"
         Height          =   195
         Left            =   180
         TabIndex        =   61
         Top             =   330
         Width           =   1125
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Fat Globules"
         Height          =   195
         Left            =   390
         TabIndex        =   60
         Top             =   1410
         Width           =   915
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "HCG Level"
         Height          =   195
         Left            =   450
         TabIndex        =   59
         Top             =   870
         Width           =   780
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Specific Gravity"
         Height          =   195
         Left            =   180
         TabIndex        =   58
         Top             =   1920
         Width           =   1125
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Demographics"
      Height          =   1935
      Index           =   1
      Left            =   150
      TabIndex        =   34
      Top             =   1560
      Width           =   9135
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
         Left            =   3930
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Copy To"
         Top             =   810
         Width           =   375
      End
      Begin VB.OptionButton cRooH 
         Caption         =   "Routine"
         Height          =   195
         Index           =   0
         Left            =   7530
         TabIndex        =   98
         Top             =   1080
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton cRooH 
         Caption         =   "Out of Hours"
         Height          =   195
         Index           =   1
         Left            =   7530
         TabIndex        =   97
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cmbComment 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5220
         TabIndex        =   45
         Top             =   540
         Width           =   3405
      End
      Begin VB.ComboBox cmbClinDetails 
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
         Height          =   315
         Left            =   5220
         Sorted          =   -1  'True
         TabIndex        =   40
         Top             =   210
         Width           =   3405
      End
      Begin VB.ComboBox cmbGP 
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
         Height          =   315
         Left            =   1170
         TabIndex        =   39
         Top             =   1140
         Width           =   2745
      End
      Begin VB.ComboBox cmbClinician 
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
         Height          =   315
         Left            =   1170
         TabIndex        =   38
         Top             =   810
         Width           =   2745
      End
      Begin VB.ComboBox cmbWard 
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
         Height          =   315
         Left            =   1170
         TabIndex        =   37
         Top             =   1470
         Width           =   2745
      End
      Begin VB.TextBox txtAddr 
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
         Index           =   1
         Left            =   1170
         MaxLength       =   40
         TabIndex        =   36
         Top             =   510
         Width           =   2745
      End
      Begin VB.TextBox txtAddr 
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
         Index           =   0
         Left            =   1170
         MaxLength       =   40
         TabIndex        =   35
         Top             =   210
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker dtCulture 
         Height          =   315
         Left            =   5640
         TabIndex        =   47
         Top             =   1290
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67305473
         CurrentDate     =   37636
      End
      Begin MSComCtl2.DTPicker dtSample 
         Height          =   315
         Left            =   5640
         TabIndex        =   48
         Top             =   960
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67305473
         CurrentDate     =   37636
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   51
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Sample Date"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   50
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Culture Date"
         Height          =   195
         Left            =   4710
         TabIndex        =   49
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Comment"
         Height          =   195
         Left            =   4470
         TabIndex        =   46
         Top             =   570
         Width           =   660
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Clin Details"
         Height          =   195
         Index           =   0
         Left            =   4350
         TabIndex        =   44
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   43
         Top             =   1530
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Clinician"
         Height          =   195
         Index           =   1
         Left            =   465
         TabIndex        =   42
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "GP"
         Height          =   195
         Index           =   1
         Left            =   825
         TabIndex        =   41
         Top             =   1200
         Width           =   225
      End
   End
   Begin VB.Frame fra 
      Height          =   1395
      Index           =   0
      Left            =   2550
      TabIndex        =   19
      Top             =   120
      Width           =   10275
      Begin VB.ComboBox cmbHospital 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   270
         TabIndex        =   33
         Text            =   "cmbHospital"
         Top             =   210
         Width           =   2085
      End
      Begin VB.CommandButton cmdSearch 
         Appearance      =   0  'Flat
         Caption         =   "Search"
         Height          =   285
         Index           =   0
         Left            =   8910
         TabIndex        =   26
         Top             =   270
         Width           =   705
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Searc&h"
         Height          =   345
         Index           =   1
         Left            =   5880
         TabIndex        =   25
         Top             =   330
         Width           =   675
      End
      Begin VB.TextBox txtSex 
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
         Left            =   7350
         MaxLength       =   6
         TabIndex        =   24
         Top             =   990
         Width           =   1545
      End
      Begin VB.TextBox txtAge 
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
         Left            =   7350
         MaxLength       =   4
         TabIndex        =   23
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtDoB 
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
         Left            =   7350
         MaxLength       =   10
         TabIndex        =   22
         Top             =   270
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
         Left            =   1890
         MaxLength       =   30
         TabIndex        =   21
         Tag             =   "tName"
         Top             =   810
         Width           =   4665
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
         Left            =   240
         MaxLength       =   8
         TabIndex        =   20
         Top             =   780
         Width           =   1545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Left            =   7020
         TabIndex        =   31
         Top             =   1020
         Width           =   270
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   0
         Left            =   6990
         TabIndex        =   30
         Top             =   660
         Width           =   285
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   0
         Left            =   6900
         TabIndex        =   29
         Top             =   300
         Width           =   405
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   2010
         TabIndex        =   28
         Top             =   570
         Width           =   420
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
         Left            =   4080
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   1755
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Colony Count"
      Height          =   5055
      Index           =   6
      Left            =   5640
      TabIndex        =   15
      Top             =   3690
      Width           =   7185
      Begin VB.ComboBox cmbCult 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4830
         TabIndex        =   92
         Text            =   "cmbCult"
         Top             =   1020
         Width           =   1815
      End
      Begin VB.ComboBox cmbCult 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   91
         Text            =   "cmbCult"
         Top             =   1020
         Width           =   1815
      End
      Begin VB.ComboBox cmbCult 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   90
         Text            =   "cmbCult"
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txtculture 
         BackColor       =   &H8000000B&
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
         Left            =   120
         MaxLength       =   40
         TabIndex        =   18
         Top             =   420
         Width           =   6975
      End
      Begin VB.ComboBox cmbCount 
         Height          =   315
         Left            =   2580
         TabIndex        =   16
         Text            =   "cmbCount"
         Top             =   120
         Width           =   1635
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   3675
         Index           =   0
         Left            =   30
         TabIndex        =   93
         Top             =   1350
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   6482
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
         GridLines       =   2
         ScrollBars      =   2
         FormatString    =   "<Antibiotic            |^RSI|^Sup"
      End
      Begin Threed.SSCommand cmdSuppress 
         Height          =   315
         Index           =   0
         Left            =   570
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   720
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Suppress"
         ForeColor       =   0
      End
      Begin Threed.SSCommand cmdSuppress 
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   95
         Top             =   720
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Suppress"
      End
      Begin Threed.SSCommand cmdSuppress 
         Height          =   315
         Index           =   2
         Left            =   5280
         TabIndex        =   96
         Top             =   720
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Suppress"
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   3675
         Index           =   1
         Left            =   2400
         TabIndex        =   120
         Top             =   1350
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   6482
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
         GridLines       =   2
         ScrollBars      =   2
         FormatString    =   "<Antibiotic            |^RSI|^Sup"
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   3675
         Index           =   2
         Left            =   4770
         TabIndex        =   121
         Top             =   1350
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   6482
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
         GridLines       =   2
         ScrollBars      =   2
         FormatString    =   "<Antibiotic            |^RSI|^Sup"
      End
      Begin VB.Image imgTop6 
         Height          =   270
         Index           =   2
         Left            =   6660
         Picture         =   "Furines.frx":2D64
         Stretch         =   -1  'True
         ToolTipText     =   "Set Top 6 Sensitive"
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgTop6 
         Height          =   270
         Index           =   1
         Left            =   4260
         Picture         =   "Furines.frx":3296
         Stretch         =   -1  'True
         ToolTipText     =   "Set Top 6 Sensitive"
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgTop6 
         Height          =   270
         Index           =   0
         Left            =   1980
         Picture         =   "Furines.frx":37C8
         Stretch         =   -1  'True
         ToolTipText     =   "Set Top 6 Sensitive"
         Top             =   1050
         Width           =   255
      End
      Begin VB.Label Label24 
         Caption         =   "/cmm"
         Height          =   195
         Left            =   4260
         TabIndex        =   17
         Top             =   150
         Width           =   405
      End
   End
   Begin VB.Frame fra 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Index           =   5
      Left            =   2940
      TabIndex        =   6
      Top             =   6360
      Width           =   2685
      Begin VB.ComboBox cmbWetPrep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   915
         TabIndex        =   10
         Top             =   790
         Width           =   1665
      End
      Begin VB.ComboBox cmbGram 
         Height          =   315
         Left            =   915
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   330
         Width           =   1665
      End
      Begin VB.TextBox txtCatalase 
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
         Left            =   915
         TabIndex        =   8
         Tag             =   "Cat"
         Top             =   1680
         Width           =   1665
      End
      Begin VB.TextBox txtCoagulase 
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
         Left            =   915
         TabIndex        =   7
         Tag             =   "Coa"
         Top             =   1250
         Width           =   1665
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Wet Prep"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   14
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Catalase"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   13
         Top             =   1710
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Coagulase"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1290
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gram Stain"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Sample Type"
      Height          =   1935
      Index           =   2
      Left            =   9390
      TabIndex        =   0
      Top             =   1560
      Width           =   3435
      Begin Threed.SSCheck chkRequest 
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   118
         Top             =   1590
         Visible         =   0   'False
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Bence Jones"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.OptionButton optSampleType 
         Caption         =   "SPA"
         Height          =   195
         Index           =   3
         Left            =   2460
         TabIndex        =   117
         Top             =   270
         Width           =   615
      End
      Begin VB.OptionButton optSampleType 
         Alignment       =   1  'Right Justify
         Caption         =   "BSU"
         Height          =   195
         Index           =   2
         Left            =   1770
         TabIndex        =   115
         Top             =   270
         Width           =   645
      End
      Begin VB.OptionButton optSampleType 
         Caption         =   "CSU"
         Height          =   195
         Index           =   1
         Left            =   990
         TabIndex        =   114
         Top             =   270
         Width           =   645
      End
      Begin VB.OptionButton optSampleType 
         Alignment       =   1  'Right Justify
         Caption         =   "MSU"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   113
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin Threed.SSCheck chkRequest 
         Height          =   195
         Index           =   5
         Left            =   1560
         TabIndex        =   1
         Top             =   1590
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Urinary HCG"
         ForeColor       =   -2147483630
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
      Begin Threed.SSCheck chkRequest 
         Height          =   195
         Index           =   4
         Left            =   1560
         TabIndex        =   2
         Top             =   1290
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Specific Gravity"
         ForeColor       =   -2147483630
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
      Begin Threed.SSCheck chkRequest 
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1290
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Fat Globules"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck chkRequest 
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   990
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Pregnancy Test"
         ForeColor       =   -2147483630
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
      Begin Threed.SSCheck chkRequest 
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   990
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "C && S"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   30
         X2              =   3450
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Test Requested"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   630
         Width           =   1140
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   150
      TabIndex        =   116
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Image imgX 
      Height          =   225
      Left            =   13260
      Picture         =   "Furines.frx":3CFA
      Top             =   3870
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmEditUrines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim loading As Boolean

Private mFromElseWHERE As Boolean
Private mFromElseWHERERunNumber As String
Private mUpdatingSensitivity As Boolean
Private mUpdatingExtra As Boolean

Private Sub CheckHistory()

Dim tb As Recordset
Dim sql As String

On Error GoTo ehCH

cmdHistory.Visible = False

sql = "SELECT TOP 1 SampleID FROM Demographics WHERE " & _
      "SampleID <> '" & Val(txtSampleID) + sysOptMicroOffset(0) & "' " & _
      "AND SampleID > '" & sysOptMicroOffset(0) & "' " & _
      "AND Chart = '" & AddTicks(txtChart) & "' " & _
      "AND PatName = '" & AddTicks(txtName) & "' "
If IsDate(txtDoB) Then
  sql = sql & "AND DoB = '" & Format$(txtDoB, "dd/MMM/yyyy") & "' "
Else
  sql = sql & "AND (DoB IS NULL OR DoB = '') "
End If

Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmdHistory.Visible = True
End If

Exit Sub

ehCH:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/CheckHistory:" & Format(er) & ":" & ers
Exit Sub

End Sub

Private Sub LoadAllDetails()

LoadDemographics
CheckHistory
LoadUrineSampleType
LoadTestRequested
LoadIsolates
LoadSensitivities
LoadUrine

cmdSave.Enabled = False

End Sub

Private Sub LoadTestRequested()

Dim tb As Recordset
Dim sql As String
Dim SampleIDWithOffset As Long
Dim n As Integer

On Error GoTo ltr

SampleIDWithOffset = Val(txtSampleID) + sysOptMicroOffset(0)

For n = 0 To 5
  chkRequest(n) = False
Next

sql = "Select Urine " & _
      "from MicroRequests where " & _
      "SampleID = '" & SampleIDWithOffset & "' "
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then

  For n = 0 To 5
    If tb!Urine And 2 ^ n Then
      chkRequest(n).Value = True
    End If
  Next

End If
Exit Sub

ltr:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/LoadTestRequested:" & Format(er) & ":" & ers
Exit Sub

End Sub

Private Function CheckConflict() As Boolean

Dim s As String
Dim Conflict As Boolean
Dim tb As Recordset
Dim sql As String
Dim Organism(0 To 2) As String
Dim n As Integer
Dim ConflictList As String
Dim ThisRunNumber As String
Dim SampleDate As String
Dim A As Integer
Dim b As Integer

On Error GoTo ehcc

CheckConflict = False
If Trim(txtChart) = "" Then
  Exit Function
End If

sql = "SELECT TOP 1 * FROM demographics WHERE " & _
      "chart = '" & Trim$(txtChart) & "' " & _
      "AND SampleDate < '" & Format(dtSample, "dd/MMM/yyyy") & "' " & _
      "AND SampleDate > '" & Format(DateAdd("d", -15, dtSample), "dd/MMM/yyyy") & "' " & _
      "AND SampleID > '" & sysOptMicroOffset(0) & "' " & _
      "ORDER BY SampleDate desc"

Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  Exit Function
End If

ThisRunNumber = tb!SampleID
SampleDate = Format(tb!SampleDate, "dd/MMM/yyyy")

sql = "SELECT DISTINCT Organism, IsolateNumber FROM Sensitivities WHERE " & _
      "SampleID = '" & ThisRunNumber & "'"
Set tb = New Recordset
RecOpenClient 0, tb, sql
If tb.EOF Then
  Exit Function
End If
Do While Not tb.EOF
  n = tb!IsolateNumber
  If n > -1 And n < 4 Then
    Organism(n) = tb!Organism & ""
  End If
  tb.MoveNext
Loop

If Trim(Organism(0) & Organism(1) & Organism(2) = "") Then
  Exit Function
End If

Conflict = False
For A = 0 To 2
  For b = 0 To 2
    If cmbCult(A) = Organism(b) Then
      Conflict = True
      Exit For
    End If
  Next
Next
If Not Conflict Then
  Exit Function
End If

Conflict = False
ConflictList = ""
For A = 0 To 2
  For b = 0 To 2
    If cmbCult(A) = Organism(b) Then
      ConflictList = SensCheck(ThisRunNumber, A, cmbCult(b))
    End If
  Next
Next

If ConflictList <> "" Then
  s = "Sensitivity Conflict" & vbCrLf & _
      "Sample Number " & ThisRunNumber & _
      " (" & Format(SampleDate, "dd/mm/yyyy") & ")" & vbCrLf & _
      ConflictList & _
      "Do you wish to procede?"
  If iMsg(s, vbYesNo + vbQuestion) = vbYes Then
    CheckConflict = False
  Else
    CheckConflict = True
  End If
End If

Exit Function

ehcc:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/CheckConflict:" & Format(er) & ":" & ers
Exit Function

End Function

Private Sub LoadUrineSampleType()

Dim tb As Recordset
Dim sql As String
Dim n As Integer

On Error GoTo ltr

sql = "Select * FROM MicroSiteDetails WHERE " & _
      "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "' " & _
      "AND Site = 'Urine'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  With tb
    Select Case tb!SiteDetails & ""
      Case "MSU": optSampleType(0).Value = True
      Case "CSU": optSampleType(1).Value = True
      Case "BSU": optSampleType(2).Value = True
      Case "SPA": optSampleType(3).Value = True
      Case Else:
        For n = 0 To 3
          optSampleType(n) = False
        Next
    End Select
  End With
Else
  optSampleType(0).Value = True
End If

Exit Sub

ltr:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/LoadUrineSampleType:" & Format(er) & ":" & ers
Exit Sub

End Sub
Private Sub SaveUrineSampleType()

Dim tb As Recordset
Dim sql As String
Dim n As Integer

On Error GoTo ltr

For n = 0 To 3
  If optSampleType(n).Value = True Then
    sql = "Select * FROM MicroSiteDetails WHERE " & _
          "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If tb.EOF Then
      tb.AddNew
      tb!SampleID = Val(txtSampleID) + sysOptMicroOffset(0)
    End If
    tb!Site = "Urine"
    Select Case n
      Case 0: tb!SiteDetails = "MSU"
      Case 1: tb!SiteDetails = "CSU"
      Case 2: tb!SiteDetails = "BSU"
      Case 3: tb!SiteDetails = "SPA"
    End Select
    tb.Update
    Exit For
  End If
Next

Exit Sub

ltr:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/LoadUrineSampleType:" & Format(er) & ":" & ers
Exit Sub

End Sub
Private Function SensCheck(ByVal SampleID As Long, _
                           ByVal Grid As Integer, _
                           ByVal Org As String) As String

Dim n As Integer
Dim CurrentRes As String
Dim OrgWas As String
Dim ConflictList As String
Dim sql As String
Dim tb As Recordset

ConflictList = ""

For n = 1 To g(Grid).Rows - 1
  If g(Grid).TextMatrix(n, 1) <> "" Then
    CurrentRes = g(Grid).TextMatrix(n, 1)
    sql = "SELECT * FROM Sensitivities WHERE " & _
          "SampleID = '" & SampleID & "' " & _
          "AND Antibiotic = '" & g(Grid).TextMatrix(n, 0) & "' " & _
          "AND Organism = '" & Org & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If Not tb.EOF Then
      OrgWas = tb!RSI & ""
      If OrgWas <> "" And (OrgWas <> CurrentRes) Then
        ConflictList = ConflictList & cmbCult(Grid) & " " & g(Grid).TextMatrix(n, 0) & " was " & _
                       Switch(OrgWas = "S", "Sensitive", _
                              OrgWas = "R", "Resistant", _
                              OrgWas = "I", "Indeterminate") & vbCrLf
      End If
    End If
  End If
Next

SensCheck = ConflictList

End Function

Private Sub SetSuppressStatus()

Dim n As Integer
Dim y As Integer

For n = 0 To 2
  cmdSuppress(n).Caption = "Suppress"
  cmdSuppress(n).ForeColor = &H0&
  cmdSuppress(n).Font.Bold = False
  g(n).Col = 2
  For y = 1 To g(n).Rows - 1
    g(n).Row = y
    If g(n).CellPicture = imgX.Picture Then
      cmdSuppress(n).Caption = "Suppressed"
      cmdSuppress(n).ForeColor = &HFF&
      cmdSuppress(n).Font.Bold = True
      Exit For
    End If
  Next
Next
  
End Sub

Private Sub SetValidStatus(ByVal V As Integer)

Dim tb As Recordset
Dim sql As String

sql = "SELECT * FROM Demographics WHERE " & _
      "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then
  tb!Valid = V
  tb.Update
End If

End Sub
Private Sub chkRequest_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub chkRequest_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

cmdSave.Enabled = True

End Sub


Private Sub cmbCasts_Click()

cmdSave.Enabled = True

End Sub

Private Sub cmbClinDetails_Click()

cmdSave.Enabled = True

End Sub


Private Sub cmbClinician_Click()

cmdSave.Enabled = True

End Sub

Private Sub cmbComment_Click()

cmdSave.Enabled = True

End Sub

Private Sub cmbComment_LostFocus()

Dim tb As Recordset
Dim sql As String

sql = "Select * FROM Lists WHERE " & _
      "ListType = 'DE' " & _
      "and Code = '" & cmbComment & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbComment = tb!Text
End If

End Sub


Private Sub cmbCount_Click()

cmdSave.Enabled = True

End Sub

Private Sub cmbCrystals_Click()

cmdSave.Enabled = True

End Sub

Private Sub cmbCult_Click(Index As Integer)

cmdSave.Enabled = True

End Sub

Private Sub cmbGP_Click()

cmdSave.Enabled = True

End Sub

Private Sub cmbGP_LostFocus()

cmbGP = QueryKnown("GP", cmbGP)

End Sub


Private Sub cmbGram_Click()

cmdSave.Enabled = True

End Sub

Private Sub cmbMisc_Click(Index As Integer)

cmdSave.Enabled = True

End Sub

Private Sub cmbWard_Click()

cmdSave.Enabled = True

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


Private Sub cmbWetPrep_Click()

cmdSave.Enabled = True

End Sub

Private Sub cmdCancel_Click()

Dim LastUsed As String

LastUsed = GetSetting("Urines", "StartUp", "LastUsed", "1")

If Val(txtSampleID) > Val(LastUsed) Then
  SaveSetting "Urines", "StartUp", "LastUsed", txtSampleID
End If

Unload Me

End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub cmdCopyTo_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub cmdHistory_Click()

With frmMicroReport
  .lblChart = txtChart
  .lblName = txtName
  .lblDoB = txtDoB
  .Show 1
End With

End Sub

Private Sub cmdHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub cmdNAD_Click()

If Trim$(txtProtein) = "" Then txtProtein = "Nil"
If Trim$(txtGlucose) = "" Then txtGlucose = "Nil"
If Trim$(txtKetones) = "" Then txtKetones = "Nil"
If Trim$(txtWCC) = "" Then txtWCC = "Nil"
If Trim$(txtRCC) = "" Then txtRCC = "Nil"
If Trim$(cmbCasts) = "" Then cmbCasts = "Nil"
If Trim$(cmbCrystals) = "" Then cmbCrystals = "Nil"
If Trim$(txtBilirubin) = "" Then txtBilirubin = "Nil"
If Trim$(txtUrobilinogen) = "" Then txtUrobilinogen = "Nil"
If Trim$(txtBloodHb) = "" Then txtBloodHb = "Nil"

End Sub

Private Sub cmdCopyTo_Click()

Dim s As String

s = cmbWard & " " & cmbClinician
s = Trim$(s) & " " & cmbGP
s = Trim$(s)

frmCopyTo.lblOriginal = s
frmCopyTo.lblSampleID = txtSampleID
frmCopyTo.Show 1

CheckCC

End Sub

Private Sub CheckCC()

Dim sql As String
Dim tb As Recordset

On Error GoTo ehCCC

cmdCopyTo.Caption = "cc"
cmdCopyTo.Font.Bold = False
cmdCopyTo.BackColor = &H8000000F

If Trim$(txtSampleID) = "" Then Exit Sub
  
sql = "Select * from SendCopyTo where " & _
      "SampleID = '" & Val(txtSampleID) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmdCopyTo.Caption = "++ cc ++"
  cmdCopyTo.Font.Bold = True
  cmdCopyTo.BackColor = &H8080FF
End If

Exit Sub

ehCCC:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/CheckCC:" & str(er) & ":" & ers
Exit Sub

End Sub

Private Sub cmdNAD_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

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

Dim PhLog As PhoneLog

PhLog = CheckPhoneLog(txtSampleID)
If PhLog.SampleID <> 0 Then
  cmdPhone.BackColor = vbYellow
  cmdPhone.Caption = "Results Phoned"
  cmdPhone.ToolTipText = "Results Phoned"
Else
  cmdPhone.BackColor = &H8000000F
  cmdPhone.Caption = "Phone Results"
  cmdPhone.ToolTipText = "Phone Results"
End If

End Sub

Private Sub cmdPhone_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub cmdPrint_Click()

cmdPrint.Visible = False

UpdateMRU Me

If cmdValidate.Caption = "&Validate" Then
  If iMsg("UnValidated Report." & vbCrLf & "OK to Print?", vbQuestion + vbYesNo) = vbNo Then
    cmdPrint.Visible = True
    Exit Sub
  End If
End If

If CheckConflict() = True Then
  cmdPrint.Visible = True
  Exit Sub
End If

SaveUrine
SaveTestRequested
SaveDemographics

If IsScreenComplete() Then
'  PrintUrine txtSampleID, True
Else
  If iMsg("Incomplete Report." & vbCrLf & "OK to Print?", vbQuestion + vbYesNo) = vbYes Then
'    PrintUrine txtSampleID, True
  End If
End If

txtSampleID = Format(Val(txtSampleID) + 1)
LoadDemographics
LoadTestRequested
LoadUrine

cmdPrint.Visible = True

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub cmdSave_Click()

If Trim$(txtName) = "" Then Exit Sub

If CheckConflict() Then
  Exit Sub
End If

SaveDemographics
SaveUrine

SaveTestRequested

SaveUrineSampleType

UpdateMRU Me

txtSampleID = Format(Val(txtSampleID) + 1)
LoadAllDetails


If txtName.Enabled Then
  txtName.SetFocus
End If

End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub cmdSearch_Click(Index As Integer)

Select Case Index

Case 0
  If Trim(txtDoB) = "" Then Exit Sub
  With frmPatHistory
    .oFor(2) = True
    .EditScreen = Me
    .FromEdit = True
    .txtName = txtDoB
    .bsearch = True
    If Not .NoPreviousDetails Then
      .Show 1
    End If
  End With
Case 1
  If Trim(txtName) = "" Then Exit Sub
  With frmPatHistory
    .oFor(0) = True
    .EditScreen = Me
    .FromEdit = True
    .txtName = txtName
    .bsearch = True
    If Not .NoPreviousDetails Then
      .Show 1
    End If
  End With
End Select

End Sub

Private Sub cmdSearch_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub cmdSetPrinter_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub cmdSuppress_Click(Index As Integer)
  
Dim y As Integer

g(Index).Col = 2

With cmdSuppress(Index)
  If .Caption = "Suppress" Then
    For y = 1 To g(Index).Rows - 1
      If g(Index).TextMatrix(y, 1) <> "" Then
        g(Index).Row = y
        Set g(Index).CellPicture = imgX.Picture
      End If
    Next
  Else
    For y = 1 To g(Index).Rows - 1
      g(Index).Row = y
      Set g(Index).CellPicture = Me.Picture
    Next
  End If
End With

SetSuppressStatus

cmdSave.Enabled = True

End Sub

Private Sub cmdSuppress_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub cmdvalidate_Click()

If Trim$(txtName) = "" Then Exit Sub

If cmdValidate.Caption = "&Validated" Then
  If iMsg("Do you want to Un-Validate?", vbQuestion + vbYesNo) = vbYes Then
    If UCase(iBOX("Password Required", , , True)) = "F1" Then
      SetValidStatus 0
    End If
  End If
  Exit Sub
End If

If Not IsScreenComplete() Then
  iMsg "Incomplete!" & vbCrLf & "Cannot Validate!", vbExclamation
  Exit Sub
End If
If CheckConflict() = True Then
  Exit Sub
End If

SaveDemographics

SaveUrine
SetValidStatus 1

SaveUrineSampleType

UpdateMRU Me

txtSampleID = Format(Val(txtSampleID) + 1)
LoadAllDetails


End Sub

Private Sub ClearUrine()
  
Dim n As Integer
Dim X As Integer

txtPregnancy = ""
txtFatGlobules = ""
txtHCGLevel = ""
txtSG = ""

txtpH = ""
txtProtein = ""
txtGlucose = ""
txtKetones = ""
txtWCC = ""
txtRCC = ""
cmbCasts = ""
cmbCrystals = ""
cmbMisc(0) = ""
cmbMisc(1) = ""
cmbMisc(2) = ""
txtBilirubin = ""
txtUrobilinogen = ""
txtBloodHb = ""
cmbCount.ListIndex = -1
cmbCount = ""

txtculture = ""

For X = 0 To 2
  For n = 0 To g(X).Rows - 1
    g(X).TextMatrix(n, 1) = ""
    cmbCult(X).ListIndex = -1
    cmbCult(X).Text = ""
  Next
Next

End Sub

Private Sub cmbCasts_LostFocus()

Dim tb As Recordset
Dim sql As String

If Trim$(cmbCasts) = "" Then
  Exit Sub
End If

sql = "Select * FROM Lists WHERE " & _
      "ListType = 'CA' " & _
      "and ( Text = '" & cmbCasts & "' " & _
      "or Code = '" & cmbCasts & "')"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbCasts = tb!Text & ""
End If

End Sub

Private Sub cmbClinician_LostFocus()

pBar = 0
cmbClinician = QueryKnown("Clin", cmbClinician, cmbHospital)

End Sub


Private Sub cmbCount_LostFocus()

Select Case UCase(cmbCount)
  Case "N": cmbCount = "Normal"
            cmdSave.SetFocus
  Case "H": cmbCount = ">100,000"
  Case "E": cmbCount = "50,000"
End Select

End Sub

Private Sub cmbCrystals_LostFocus()

Dim tb As Recordset
Dim sql As String

If Trim$(cmbCrystals) = "" Then
  Exit Sub
End If

sql = "Select * FROM Lists WHERE " & _
      "ListType = 'CR' " & _
      "and ( Text = '" & cmbCrystals & "' " & _
      "or Code = '" & cmbCrystals & "')"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbCrystals = tb!Text & ""
End If

End Sub


Private Sub cmbCult_LostFocus(Index As Integer)

Dim n As Integer
Dim y As Integer
Dim tb As Recordset
Dim sql As String

If Trim$(cmbCasts) = "" Then
  Exit Sub
End If

sql = "Select * FROM Lists WHERE " & _
      "ListType = 'OR' " & _
      "and ( Text = '" & cmbCult(Index) & "' " & _
      "or Code = '" & cmbCult(Index) & "')"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
   cmbCult(Index) = tb!Text & ""
End If

txtculture = ""
If cmbCult(0) <> "" Then
  txtculture = cmbCult(0)
End If
If cmbCult(1) <> "" Then
  If txtculture <> "" Then
    txtculture = txtculture & " & "
  End If
  txtculture = txtculture & cmbCult(1)
End If
If cmbCult(2) <> "" Then
  If txtculture <> "" Then
    txtculture = txtculture & " & "
  End If
  txtculture = txtculture & cmbCult(2)
End If

If Index = 0 Then
  If Trim(UCase(cmbCult(0))) = "MIXED GROWTH" Then
    cmdSave.SetFocus
  End If
End If

'Check Proteus and Nitrofuratoin
For n = 0 To 2
  If Left(UCase(cmbCult(n)), 7) = "PROTEUS" Then
    g(n).Col = 0
    For y = 0 To g(n).Rows - 1
      g(n).Row = y
      If g(n) = "Nitrofuratoin" Then
        g(n).Col = 1
        g(n) = "R"
        Exit For
      End If
    Next
  End If
Next

End Sub


Private Sub cmbGram_LostFocus()

Dim tb As Recordset
Dim sql As String

If Trim$(cmbGram) = "" Then
  Exit Sub
End If

sql = "Select * FROM Lists WHERE " & _
      "ListType = 'GS' " & _
      "and ( Text = '" & cmbGram & "' " & _
      "or Code = '" & cmbGram & "')"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbGram = tb!Text & ""
End If

End Sub


Private Sub cmbMisc_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String

If Trim$(cmbMisc(Index)) = "" Then
  Exit Sub
End If

sql = "Select * FROM Lists WHERE " & _
      "ListType = 'MI' " & _
      "and ( Text = '" & cmbMisc(Index) & "' " & _
      "or Code = '" & cmbMisc(Index) & "')"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbMisc(Index) = tb!Text & ""
End If

End Sub



Private Sub FillPathogens()

Dim tb As Recordset
Dim sql As String

cmbCult(0).Clear
cmbCult(1).Clear
cmbCult(2).Clear

sql = "SELECT * FROM Lists WHERE " & _
      "ListType = 'OR' " & _
      "ORDER BY ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbCult(0).AddItem tb!Text
  cmbCult(1).AddItem tb!Text
  cmbCult(2).AddItem tb!Text
  tb.MoveNext
Loop

End Sub

Private Sub cmbWetPrep_LostFocus()

Dim tb As Recordset
Dim sql As String

If Trim$(cmbWetPrep) = "" Then
  Exit Sub
End If

sql = "Select * FROM Lists WHERE " & _
      "ListType = 'WP' " & _
      "and ( Text = '" & cmbWetPrep & "' " & _
      "or Code = '" & cmbWetPrep & "')"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  cmbWetPrep = tb!Text & ""
End If

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

End Sub


Private Sub cmdValidate_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub cMRU_Click()

txtSampleID = cMRU

LoadAllDetails

End Sub


Private Sub cRooH_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub cRooH_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

cmdSave.Enabled = True

End Sub


Private Sub dtCulture_CloseUp()

cmdSave.Enabled = True

End Sub

Private Sub dtCulture_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub dtSample_CloseUp()

cmdSave.Enabled = True

End Sub

Private Sub dtSample_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub Form_Load()

Dim n As Integer

udSampleID.max = 9999999

FillMRU Me

dtSample = Format(Now, "dd/mmm/yyyy")
dtCulture = Format(Now, "dd/mmm/yyyy")

FillLists
FillAbGrid
FillPathogens
FillCasts
FillCrystals
FillMiscellaneous

cmbCount.Clear
For n = 1 To 14
  cmbCount.AddItem Choose(n, "Normal", ">100,000", "100,000", "90,000", _
      "80,000", "70,000", "60,000", "50,000", _
      "40,000", "30,000", "20,000", "15,000", _
      "10,000", "5,000")
Next

LoadInitial

Screen.MousePointer = 0

End Sub

Private Function IsScreenComplete() As Boolean
  
Dim X As Integer
Dim y As Integer

IsScreenComplete = True
If InStr(UCase(cmbCult(0)), "MIXED GROWTH") <> 0 Then Exit Function
If InStr(UCase(cmbCult(1)), "MIXED GROWTH") <> 0 Then Exit Function
If InStr(UCase(cmbCult(2)), "MIXED GROWTH") <> 0 Then Exit Function

If UCase(cmbGram) = "PENDING" _
   Or UCase(cmbWetPrep) = "PENDING" _
   Or UCase(txtCoagulase) = "PENDING" _
   Or UCase(txtCatalase) = "PENDING" Then
  IsScreenComplete = False
  Exit Function
End If

If chkRequest(1) And Trim(txtPregnancy) = "" _
   Or chkRequest(2) And Trim(txtPregnancy) = "" _
   Or chkRequest(3) And Trim(txtFatGlobules) = "" _
   Or chkRequest(4) And Trim(txtSG) = "" _
   Or chkRequest(5) And Trim(txtHCGLevel) = "" _
   Or txtProtein = "Pos" _
   Or txtGlucose = "Pos" Then
  IsScreenComplete = False
  Exit Function
End If
  
If (UCase(cmbCult(0)) <> "CANDIDA" _
   And InStr(UCase(cmbCult(0)), "MIXED GROWTH") <> 0 _
   And Trim(cmbCult(0)) <> "" _
   And UCase(cmbCult(0)) <> "STERILE") _
   Or _
   (UCase(cmbCult(1)) <> "CANDIDA" _
   And InStr(UCase(cmbCult(1)), "MIXED GROWTH") <> 0 _
   And Trim(cmbCult(1)) <> "" _
   And UCase(cmbCult(1)) <> "STERILE") _
   Or _
   (UCase(cmbCult(2)) <> "CANDIDA" _
   And InStr(UCase(cmbCult(2)), "MIXED GROWTH") <> 0 _
   And Trim(cmbCult(2)) <> "" _
   And UCase(cmbCult(2)) <> "STERILE") Then
    
  IsScreenComplete = False
  For X = 0 To 2
    g(X).Col = 1
    For y = 1 To g(X).Rows - 1
      g(X).Row = y
      If g(X) <> "" Then
        IsScreenComplete = True
        Exit Function
      End If
    Next
  Next
End If

End Function



Private Sub SaveTestRequested()

Dim lngU As Long
Dim n As Integer
Dim tb As Recordset
Dim sql As String
Dim SampleIDWithOffset As Long

lngU = 0
For n = 0 To 5
  If chkRequest(n) Then
    lngU = lngU + 2 ^ n
  End If
Next

SampleIDWithOffset = Val(txtSampleID) + sysOptMicroOffset(0)

sql = "Select * from MicroRequests where " & _
      "SampleID = '" & SampleIDWithOffset & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
End If
tb!SampleID = SampleIDWithOffset
tb!RequestDate = Format(Now, "dd/mmm/yyyy hh:mm")
tb!Faecal = 0
tb!Urine = lngU
tb.Update

Exit Sub

str:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/SaveTestRequested:" & Format(er) & ":" & ers
Exit Sub

End Sub

Sub FillCrystals()

Dim tb As Recordset
Dim sql As String

cmbCrystals.Clear

sql = "Select * FROM Lists WHERE " & _
      "ListType = 'CR'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbCrystals.AddItem tb!Text & ""
  tb.MoveNext
Loop
cmbCrystals.AddItem "", 0

End Sub


Sub FillCasts()

Dim tb As Recordset
Dim sql As String

cmbCrystals.Clear

sql = "Select * FROM Lists WHERE " & _
      "ListType = 'CA'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbCasts.AddItem tb!Text & ""
  tb.MoveNext
Loop
cmbCasts.AddItem "", 0

End Sub

Sub FillMiscellaneous()

Dim tb As Recordset
Dim sql As String

cmbMisc(0).Clear
cmbMisc(1).Clear
cmbMisc(2).Clear

sql = "Select * FROM Lists WHERE " & _
      "ListType = 'MI'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbMisc(0).AddItem tb!Text & ""
  cmbMisc(1).AddItem tb!Text & ""
  cmbMisc(2).AddItem tb!Text & ""
  tb.MoveNext
Loop
cmbMisc(0).AddItem "Nil", 0
cmbMisc(1).AddItem "Nil", 0
cmbMisc(2).AddItem "Nil", 0

End Sub

Sub FillLists()

Dim tb As Recordset
Dim sql As String
Dim HospCode As String

HospCode = ListCodeFor("HO", cmbHospital)

cmbWard.Clear
sql = "Select * FROM Wards WHERE " & _
      "HospitalCode = '" & HospCode & "' " & _
      "Order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbWard.AddItem tb!Text & ""
  tb.MoveNext
Loop

cmbClinician.Clear
sql = "Select * FROM Clinicians WHERE " & _
      "HospitalCode = '" & HospCode & "' " & _
      "Order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbClinician.AddItem tb!Text & ""
  tb.MoveNext
Loop

cmbGP.Clear
sql = "Select * FROM GPs WHERE " & _
      "HospitalCode = '" & HospCode & "' " & _
      "Order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbGP.AddItem tb!Text & ""
  tb.MoveNext
Loop

cmbClinDetails.Clear
sql = "Select * FROM Lists WHERE " & _
      "Code = 'UR' " & _
      "Order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbClinDetails.AddItem tb!Text & ""
  tb.MoveNext
Loop

cmbComment.Clear
sql = "Select * FROM Lists WHERE " & _
      "Code = 'UC' " & _
      "Order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbComment.AddItem tb!Text & ""
  tb.MoveNext
Loop

cmbGram.Clear
sql = "Select * FROM Lists WHERE " & _
      "Code = 'GS' " & _
      "Order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbGram.AddItem tb!Text & ""
  tb.MoveNext
Loop

cmbWetPrep.Clear
sql = "Select * FROM Lists WHERE " & _
      "Code = 'WP' " & _
      "Order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  cmbWetPrep.AddItem tb!Text & ""
  tb.MoveNext
Loop

End Sub


Private Sub LoadInitial()

If mFromElseWHERE Then
  txtSampleID = mFromElseWHERERunNumber
Else
  txtSampleID = GetSetting("Urines", "StartUp", "LastUsed", "1")
End If

LoadAllDetails

End Sub

Private Sub LoadIsolates()

Dim tb As Recordset
Dim sql As String
Dim Iso As Integer

cmbCult(0).Text = ""
cmbCult(1).Text = ""
cmbCult(2).Text = ""

sql = "Select * FROM Isolates WHERE " & _
      "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "' and IsolateNumber < 3"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  Iso = tb!IsolateNumber
  cmbCult(Iso).Text = tb!OrganismName & ""
  tb.MoveNext
Loop

End Sub

Private Sub SaveIsolates()

Dim tb As Recordset
Dim sql As String
Dim Iso As Integer

For Iso = 0 To 2
  sql = "Select * FROM Isolates WHERE " & _
        "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "' " & _
        "AND IsolateNumber = '" & Iso & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If Trim$(cmbCult(Iso).Text) = "" Then
    If Not tb.EOF Then
      tb.Delete
    End If
  Else
    If tb.EOF Then
      tb.AddNew
      tb!SampleID = Val(txtSampleID) + sysOptMicroOffset(0)
    End If
    tb!IsolateNumber = Iso
    tb!OrganismName = cmbCult(Iso).Text
    tb.Update
  End If
Next

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim SensDone As Boolean

mFromElseWHERE = False

If mUpdatingSensitivity Then
  g(0).Col = 1
  SensDone = True
  For n = 0 To 4
    g(0).Row = n
    If Trim(g(0)) = "" Then
      SensDone = False
    End If
  Next
  If SensDone Then
    sql = "SELECT * FROM UrineIdent WHERE " & _
          "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If Not tb.EOF Then
      tb!urinesensitivity = "Done"
      tb.Update
    End If
  End If
End If

SensDone = True
If mUpdatingExtra Then
  For n = 5 To g(0).Rows - 1
    g(0).Row = n
    If Trim(g(0)) = "" Then
      SensDone = False
    End If
  Next
  If SensDone Then
    sql = "SELECT * FROM UrineIdent WHERE " & _
          "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If Not tb.EOF Then
      tb!extrasensitivity = "Done"
      tb.Update
    End If
  End If
End If

mUpdatingSensitivity = False
mUpdatingExtra = False

End Sub

Private Sub fra_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub fraSampleID_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub FillAbGrid()

Dim n As Integer
Dim tb As Recordset
Dim sql As String

On Error GoTo ehla

For n = 0 To 2
  g(n).Visible = False
  g(n).Rows = 2
  g(n).AddItem ""
  g(n).RemoveItem 1
Next

sql = "SELECT AntibioticName, MAX(ListOrder) AS M FROM Antibiotics " & _
      "GROUP BY AntibioticName ORDER BY M"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  g(0).AddItem Trim$(tb!AntibioticName & "")
  g(1).AddItem Trim$(tb!AntibioticName & "")
  g(2).AddItem Trim$(tb!AntibioticName & "")
  tb.MoveNext
Loop

For n = 0 To 2
  If g(n).Rows > 2 Then
    g(n).RemoveItem 1
  End If
  g(n).Visible = True
Next

Exit Sub

ehla:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
g(0).Visible = True
g(1).Visible = True
LogError "frmEditUrines/FillAbGrid:" & Format(er) & ":" & ers
Exit Sub

End Sub

Private Sub LoadSensitivities()

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim y As Integer

On Error GoTo ehls

FillAbGrid 'This clears the grids

For n = 0 To 2

  sql = "SELECT S.RSI, S.Report, A.AntibioticName FROM Sensitivities AS S, Antibiotics as A WHERE " & _
        "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "' " & _
        "AND A.AntibioticName = S.Antibiotic " & _
        "AND IsolateNumber = '" & n & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  Do While Not tb.EOF
    For y = 1 To g(n).Rows - 1
      If g(n).TextMatrix(y, 0) = Trim$(tb!AntibioticName) Then
        g(n).TextMatrix(y, 1) = tb!RSI & ""
        g(n).Row = y
        g(n).Col = 2
        If tb!Report = 0 Then
          Set g(n).CellPicture = imgX.Picture
        Else
          Set g(n).CellPicture = Me.Picture
        End If
        Exit For
      End If
    Next
    tb.MoveNext
  Loop

Next
SetSuppressStatus

Exit Sub

ehls:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
g(0).Visible = True
g(1).Visible = True
g(2).Visible = True
LogError "frmEditUrines/LoadSensitivities:" & Format(er) & ":" & ers
Exit Sub

End Sub

Private Sub LoadUrine()

Dim tb As Recordset
Dim sql As String

On Error GoTo ehlu

loading = True

sql = "Select * FROM Urine WHERE " & _
      "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  ClearUrine
Else
  Select Case UCase$(Left$(tb!Pregnancy & "", 1))
    Case "I": txtPregnancy = "Inconclusive"
    Case "N": txtPregnancy = "Negative"
    Case "P": txtPregnancy = "Positive"
    Case "U": txtPregnancy = "Unsuitable"
    Case Else: txtPregnancy = ""
  End Select
  txtFatGlobules = tb!FatGlobules & ""
  txtHCGLevel = tb!HCGLevel & ""
  txtSG = tb!SG & ""
  txtpH = tb!pH & ""
  txtProtein = tb!Protein & ""
  txtGlucose = tb!Glucose & ""
  txtKetones = tb!Ketones & ""
  txtWCC = tb!WCC & ""
  txtRCC = tb!RCC & ""
  cmbCasts = tb!Casts & ""
  cmbCrystals = tb!Crystals & ""
  txtBilirubin = tb!Bilirubin & ""
  txtUrobilinogen = tb!Urobilinogen & ""
  txtBloodHb = tb!BloodHb & ""
  cmbMisc(0) = tb!Misc0 & ""
  cmbMisc(1) = tb!Misc1 & ""
  cmbMisc(2) = tb!Misc2 & ""
  cmbCount = tb!Count & ""
  LoadIsolates
  txtculture = Trim$(cmbCult(0) & " " & cmbCult(1) & " " & cmbCult(2))
End If

LoadSensitivities

sql = "SELECT * FROM UrineIdent WHERE " & _
      "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  cmbGram = ""
  cmbWetPrep = ""
  txtCoagulase = ""
  txtCatalase = ""
Else
  cmbGram = tb!Gram & ""
  cmbWetPrep = tb!WetPrep & ""
  txtCoagulase = tb!Coagulase & ""
  txtCatalase = tb!Catalase & ""
End If

If txtChart.Enabled And txtChart.Visible = True And Trim(txtChart) = "" Then
  txtChart.SetFocus
ElseIf txtpH.Enabled And txtpH.Visible Then
  txtpH.SetFocus
ElseIf cmbCount.Enabled And cmbCount.Visible Then
  cmbCount.SetFocus
Else
  If cmdSave.Enabled And cmdSave.Visible Then
    cmdSave.SetFocus
  End If
End If

loading = False

Exit Sub

ehlu:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/LoadUrine:" & Format(er) & ":" & ers
Exit Sub

End Sub
Private Sub SaveSensitivities()

Dim graf As Integer
Dim n As Integer
Dim Report As Integer
Dim Code As String
Dim sql As String

On Error GoTo ehss

Screen.MousePointer = vbHourglass

sql = "DELETE FROM Sensitivities WHERE " & _
      "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
Cnxn(0).Execute sql

For graf = 0 To 2
  For n = 1 To g(graf).Rows - 1
    If g(graf).TextMatrix(n, 1) <> "" Then
      Code = AntibioticCodeFor(g(graf).TextMatrix(n, 0))
      Debug.Print Code
      If Code <> "???" Then
        g(graf).Row = n
        g(graf).Col = 2
        Report = IIf(g(graf).CellPicture = imgX.Picture, 0, 1)
        
        sql = "INSERT INTO Sensitivities " & _
              "(SampleID, Organism, Antibiotic, AntibioticCode, RSI, IsolateNumber, Report ) VALUES (" & _
              "'" & Val(txtSampleID) + sysOptMicroOffset(0) & "', " & _
              "'" & cmbCult(graf) & "', " & _
              "'" & g(graf).TextMatrix(n, 0) & "', " & _
              "'" & Code & "', " & _
              "'" & g(graf).TextMatrix(n, 1) & "', " & _
              "'" & graf & "', " & _
              "'" & Report & "')"
        Debug.Print sql
        
        Cnxn(0).Execute sql
      End If
    End If
  Next
Next

Screen.MousePointer = 0

Exit Sub

ehss:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/SaveSensitivities:" & Format(er) & ":" & ers
Exit Sub

End Sub

Private Sub SaveUrine()

Dim tb As Recordset
Dim sql As String
Dim Saveable As Boolean

On Error GoTo ehsu

Screen.MousePointer = vbHourglass

sql = "Select * FROM Urine WHERE " & _
      "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
  tb!SampleID = Val(txtSampleID) + sysOptMicroOffset(0)
End If

tb!Pregnancy = Left$(txtPregnancy, 1)
tb!FatGlobules = txtFatGlobules
tb!HCGLevel = txtHCGLevel
tb!SG = txtSG
tb!pH = txtpH
tb!Protein = txtProtein
tb!Glucose = txtGlucose
tb!Ketones = txtKetones
tb!WCC = txtWCC
tb!RCC = txtRCC
tb!Casts = Left(Trim(cmbCasts), 20)
tb!Crystals = Left(Trim(cmbCrystals), 20)
tb!Bilirubin = txtBilirubin
tb!Urobilinogen = txtUrobilinogen
tb!BloodHb = txtBloodHb
tb!Misc0 = Left(Trim(cmbMisc(0)), 50)
tb!Misc1 = Left(Trim(cmbMisc(1)), 50)
tb!Misc2 = Left(Trim(cmbMisc(2)), 50)
tb!Count = cmbCount
SaveIsolates

'tb!CultureDate = Format(dtCulture, "dd/mmm/yyyy")
tb.Update

Saveable = False
If Trim(cmbGram & cmbWetPrep & txtCoagulase & txtCatalase) <> "" Then
  Saveable = True
End If

sql = "SELECT * FROM UrineIdent WHERE " & _
      "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  If Saveable Then
    tb.AddNew
    tb!SampleID = Val(txtSampleID) + sysOptMicroOffset(0)
  End If
Else
  If Not Saveable Then
    tb.Delete
  End If
End If
If Saveable Then
  tb!Gram = cmbGram
  tb!WetPrep = cmbWetPrep
  tb!Coagulase = txtCoagulase
  tb!Catalase = txtCatalase
  tb.Update
End If

SaveSensitivities

Screen.MousePointer = 0

Exit Sub

ehsu:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/SaveUrine:" & Format(er) & ":" & ers
Exit Sub

End Sub

Private Sub g_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub g_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

On Error GoTo ehgc

If g(Index).MouseRow = 0 Then Exit Sub
If Button <> vbLeftButton Then Exit Sub

With g(Index)
  
  If Left(UCase(cmbCult(Index)), 7) = "PROTEUS" Then
    If .TextMatrix(.Row, 0) = "Nitrofuratoin" Then
      .TextMatrix(.Row, 1) = "R"
    End If
  End If
  
  If .Col = 1 Then
    .TextMatrix(.Row, 1) = "R"
  ElseIf .Col = 0 Then
    Select Case .TextMatrix(.Row, 1)
      Case "":
        .TextMatrix(.Row, 1) = "S"
      Case "S":
        .TextMatrix(.Row, 1) = "R"
      Case "R":
        .TextMatrix(.Row, 1) = "I"
      Case "I":
        .TextMatrix(.Row, 1) = ""
        .Col = 2
        Set .CellPicture = Me.Picture
        .Col = 1
        SetSuppressStatus
    End Select
  Else
    If .TextMatrix(.Row, 1) <> "" Then
      If .CellPicture = imgX.Picture Then
        Set .CellPicture = Me.Picture
      Else
        Set .CellPicture = imgX.Picture
      End If
      SetSuppressStatus
    End If
  End If
  
End With

cmdSave.Enabled = True
cmdSave.SetFocus

Exit Sub

ehgc:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/g_Click:" & Format(er) & ":" & ers

Exit Sub

End Sub

Private Sub imgLast_Click()

Dim tb As Recordset
Dim sql As String

sql = "SELECT TOP 1 SampleID FROM Demographics WHERE " & _
      "SampleID > " & sysOptMicroOffset(0) & " " & _
      "AND SampleID < " & sysOptMicroOffset(0) + 10000000 & "  " & _
      "Order by SampleID desc"

Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  txtSampleID = Format$(tb!SampleID - sysOptMicroOffset(0))
Else
  txtSampleID = "1"
End If
LoadAllDetails

End Sub

Private Sub imgLast_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub imgTop6_Click(Index As Integer)

Dim n As Integer

For n = 1 To 6
  If g(Index).Rows > n Then
    g(Index).TextMatrix(n, 1) = "S"
  End If
Next
cmdSave.Enabled = True

End Sub

Private Sub cmbHospital_Click()

FillLists

cmdSave.Enabled = True

End Sub

Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub imgTop6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub irelevant_Click(Index As Integer)

Dim tb As Recordset
Dim sql As String
Dim Direction As String

If Index = 0 Then
  Direction = "<"
Else
  Direction = ">"
End If

sql = "SELECT TOP 1 SampleID FROM Demographics WHERE " & _
      "SampleID > " & sysOptMicroOffset(0) & "  " & _
      "AND SampleID " & Direction & " " & Val(txtSampleID) + sysOptMicroOffset(0) & " " & _
      "Order by SampleID " & IIf(Direction = "<", "desc", "asc")

Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  txtSampleID = Format$(tb!SampleID - sysOptMicroOffset(0))
Else
  txtSampleID = "1"
End If
LoadAllDetails

End Sub

Private Sub iRelevant_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub optSampleType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

cmdSave.Enabled = True

pBar = 0

End Sub


Private Sub pBar_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub TimerBar_Timer()

pBar = pBar + 1

If pBar >= pBar.max Then
  Unload Me
End If

End Sub

Private Sub txtAddr_KeyPress(Index As Integer, KeyAscii As Integer)

cmdSave.Enabled = True

End Sub


Private Sub txtAddr_LostFocus(Index As Integer)

txtAddr(Index) = Initial2Upper(txtAddr(Index))

End Sub


Private Sub txtAddr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtAge_Click()

cmdSave.Enabled = True

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

cmdSave.Enabled = True

End Sub


Private Sub txtAge_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

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

cmdSave.Enabled = True

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

cmdSave.Enabled = True

End Sub


Private Sub txtBilirubin_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

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

cmdSave.Enabled = True

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

cmdSave.Enabled = True

End Sub


Private Sub txtBloodHb_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub txtCatalase_Click()

ClickMe txtCatalase

End Sub


Private Sub txtCatalase_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtChart_KeyPress(KeyAscii As Integer)

cmdSave.Enabled = True

End Sub

Private Sub txtchart_LostFocus()

Dim tb As Recordset
Dim sql As String

On Error GoTo ehclf

If Trim(txtChart) = "" Then Exit Sub

Screen.MousePointer = vbHourglass

sql = "SELECT TOP 1 * FROM Demographics WHERE " & _
      "Chart = '" & txtChart & "' " & _
      "ORDER BY SampleDate desc"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  txtName = tb!PatName & ""
  txtDoB = tb!DoB & ""
  txtAge = tb!Age & ""
  txtSex = tb!Sex & ""
  txtAddr(0) = tb!Addr0 & ""
  txtAddr(1) = tb!Addr1 & ""
  cmbWard = tb!Ward & ""
  cmbClinician = tb!Clinician & ""
  cmbGP = tb!GP & ""
Else
  txtName = ""
  txtDoB = ""
  txtAge = ""
  txtSex = ""
  txtAddr(0) = ""
  txtAddr(1) = ""
  cmbWard = ""
  cmbClinician = ""
  cmbGP = ""
End If

Screen.MousePointer = 0

Exit Sub

ehclf:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/txtChart_LostFocus:" & Format(er) & ":" & ers
Exit Sub

End Sub


Private Sub txtChart_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtCoagulase_Click()

ClickMe txtCoagulase

End Sub


Private Sub txtCoagulase_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtculture_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtDoB_KeyPress(KeyAscii As Integer)

cmdSave.Enabled = True

End Sub


Private Sub txtDoB_LostFocus()

txtDoB = Convert62Date(txtDoB, BACKWARD)
txtAge = CalcAge(txtDoB)

End Sub

Private Sub txtDoB_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtFatGlobules_Click()

Select Case txtFatGlobules
  Case "": txtFatGlobules = "Not Seen"
  Case "Not Seen": txtFatGlobules = "Present"
  Case Else: txtFatGlobules = ""
End Select

cmdSave.Enabled = True

End Sub


Private Sub txtFatGlobules_KeyPress(KeyAscii As Integer)

KeyAscii = 0

Select Case txtFatGlobules
  Case "": txtFatGlobules = "Not Seen"
  Case "Not Seen": txtFatGlobules = "Present"
  Case Else: txtFatGlobules = ""
End Select

cmdSave.Enabled = True

End Sub


Private Sub txtFatGlobules_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub txtGlucose_Click()

If txtGlucose = "" Then
  txtGlucose = "Pos"
Else
  txtGlucose.SelStart = 0
  txtGlucose.SelLength = Len(txtGlucose)
End If

cmdSave.Enabled = True

End Sub


Private Sub txtGlucose_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtHCGLevel_KeyPress(KeyAscii As Integer)

cmdSave.Enabled = True

End Sub

Private Sub txtHCGLevel_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

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

cmdSave.Enabled = True

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

cmdSave.Enabled = True

End Sub



Private Sub txtKetones_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

cmdSave.Enabled = True

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



Private Sub txtName_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtpH_Click()

Select Case txtpH
  Case "": txtpH = "Acid"
  Case "Acid": txtpH = "Alkaline"
  Case "Alkaline": txtpH = "Neutral"
  
  Case "Neutral":
    If iMsg("Is Sample Unsuitable?", vbQuestion + vbYesNo) = vbYes Then
      txtpH = "Unsuitable"
      txtProtein = ""
      txtGlucose = ""
      txtKetones = ""
      txtUrobilinogen = ""
      txtBilirubin = ""
    Else
      txtpH = ""
    End If
    
  Case Else: txtpH = ""
End Select

cmdSave.Enabled = True

End Sub

Private Sub txtPh_KeyPress(KeyAscii As Integer)

KeyAscii = 0

Select Case txtpH
  Case "": txtpH = "Acid"
  Case "Acid": txtpH = "Alkaline"
  Case "Alkaline": txtpH = "Neutral"
  Case "Neutral": txtpH = "Unsuitable"
  Case Else: txtpH = ""
End Select

cmdSave.Enabled = True

End Sub


Private Sub txtpH_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub txtPregnancy_KeyPress(KeyAscii As Integer)

Select Case UCase(Chr(KeyAscii))
  Case "N": txtPregnancy = "Negative"
  Case "P": txtPregnancy = "Positive"
  Case "I": txtPregnancy = "Inconclusive"
  Case "U": txtPregnancy = "Unsuitable"
  Case Else: txtPregnancy = ""
End Select

KeyAscii = 0

cmdSave.Enabled = True

End Sub


Private Sub txtPregnancy_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtProtein_Click()

If txtProtein = "" Then
  txtProtein = "Pos"
Else
  txtProtein.SelStart = 0
  txtProtein.SelLength = Len(txtProtein)
End If

CheckIfDoSensitivity

cmdSave.Enabled = True

End Sub


Private Sub txtProtein_KeyPress(KeyAscii As Integer)

CheckIfDoSensitivity

cmdSave.Enabled = True

End Sub


Private Sub txtProtein_LostFocus()

CheckIfDoSensitivity

End Sub


Private Sub txtProtein_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub txtRCC_KeyPress(KeyAscii As Integer)

Select Case Chr(KeyAscii)
  Case "P", "p":
    txtRCC = "Packed"
    KeyAscii = 0
  Case ">", "G", "g":
    txtRCC = ">200"
    KeyAscii = 0
  Case "N", "n":
    txtRCC = "Nil"
    KeyAscii = 0
End Select

cmdSave.Enabled = True

End Sub


Private Sub txtRCC_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtSampleID_LostFocus()

Dim n As Long

n = Val(txtSampleID)
If n > 9999999 Then
  iMsg "Incorrect Sample Number Format", vbExclamation
  n = 1
End If

txtSampleID = Format(n)

LoadAllDetails

If txtName.Enabled Then
  txtName.SetFocus
End If

End Sub
Private Sub LoadDemographics()

Dim tb As Recordset
Dim sql As String
Dim Cx As Comment
Dim Cxs As New Comments
Dim n As Integer

On Error GoTo ehlud

cmdValidate.Caption = "&Validate"
For n = 0 To 6
  fra(n).Enabled = True
Next

sql = "Select * FROM Demographics WHERE " & _
      "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  cRooH(0).Value = True
  cmbHospital = HospName(0)
  cmdPrint.Caption = "&Print"
  dtSample = Format(Now, "dd/mmm/yyyy")
  txtChart = ""
  txtName = ""
  txtDoB = ""
  txtAge = ""
  txtSex = ""
  txtAddr(0) = ""
  txtAddr(1) = ""
  cmbWard = ""
  cmbClinician = ""
  cmbGP = ""
  cmbClinDetails = ""
  cmbComment = ""
Else
  If Trim$(tb!Hospital & "") = "" Then
    cmbHospital = HospName(0)
  Else
    cmbHospital = tb!Hospital
  End If
'  cmdPrint.Caption = IIf(tb!urineprinted, "Re&print", "&Print")
  If Not IsNull(tb!RooH) Then
    cRooH(0).Value = tb!RooH
    cRooH(1).Value = Not cRooH(0).Value
  End If
  
  dtSample = Format(tb!SampleDate, "dd/mmm/yyyy")
  txtChart = Trim$(tb!Chart & "")
  txtName = tb!PatName & ""
  txtDoB = tb!DoB & ""
  txtAge = tb!Age & ""
  txtSex = tb!Sex & ""
  txtAddr(0) = tb!Addr0 & ""
  txtAddr(1) = tb!Addr1 & ""
  cmbWard = tb!Ward & ""
  cmbClinician = tb!Clinician & ""
  cmbGP = tb!GP & ""
  cmbClinDetails = tb!ClDetails & ""
  Set Cx = Cxs.Load(Val(txtSampleID) + sysOptMicroOffset(0))
  If Not Cx Is Nothing Then
    cmbComment = Cx.MicroGeneral
  Else
    cmbComment = ""
  End If
  If Not IsNull(tb!Rundate) Then
    dtCulture = Format(tb!Rundate, "dd/mmm/yyyy")
  Else
    dtCulture = Format(Now, "dd/mmm/yyyy")
  End If
  If Not IsNull(tb!Valid) Then
    If tb!Valid Then
      cmdValidate.Caption = "&Validated"
      For n = 0 To 6
        fra(n).Enabled = False
      Next
    End If
  End If
End If

CheckIfPhoned

Exit Sub

ehlud:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/LoadDemographics:" & Format(er) & ":" & ers
Exit Sub

End Sub

Private Sub SaveDemographics()

Dim tb As Recordset
Dim sql As String
Dim Cx As New Comment
Dim Cxs As New Comments

On Error GoTo ehSD

sql = "Select * FROM Demographics WHERE " & _
      "SampleID = '" & Val(txtSampleID) + sysOptMicroOffset(0) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
End If

tb!SampleID = Val(txtSampleID) + sysOptMicroOffset(0)
tb!RooH = cRooH(0).Value
tb!Hospital = cmbHospital
  
tb!SampleDate = Format(dtSample, "dd/MMM/yyyy")
tb!Chart = AddTicks(txtChart)
tb!PatName = AddTicks(txtName)
If IsDate(txtDoB) Then
  tb!DoB = Format$(txtDoB, "dd/mmm/yyyy")
Else
  tb!DoB = Null
End If
tb!Age = txtAge
tb!Sex = Left$(txtSex, 1)
tb!Addr0 = AddTicks(txtAddr(0))
tb!Addr1 = AddTicks(txtAddr(1))
tb!Ward = AddTicks(cmbWard)
tb!Clinician = AddTicks(cmbClinician)
tb!GP = AddTicks(cmbGP)
tb!ClDetails = AddTicks(cmbClinDetails)
tb!Rundate = Format$(dtCulture, "dd/MMM/yyyy")
tb.Update

Cx.SampleID = Val(txtSampleID) + sysOptMicroOffset(0)
Cx.MicroGeneral = AddTicks(cmbComment)
Cxs.Save Cx
  
Exit Sub

ehSD:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmEditUrines/SaveDemographics:" & Format(er) & ":" & ers
Exit Sub

End Sub

Private Sub txtSampleID_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub


Private Sub txtsex_Click()

Select Case Left$(txtSex, 1)
  Case "": txtSex = "Male"
  Case "M": txtSex = "Female"
  Case Else: txtSex = ""
End Select

cmdSave.Enabled = True

End Sub

Private Sub txtsex_KeyPress(KeyAscii As Integer)

Select Case Trim(txtSex)
  Case "": txtSex = "M"
  Case "M": txtSex = "F"
  Case Else: txtSex = ""
End Select

KeyAscii = 0

cmdSave.Enabled = True

End Sub


Private Sub txtSex_LostFocus()

SexLostFocus txtSex, txtName

End Sub

Private Sub txtSex_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub txtSG_KeyPress(KeyAscii As Integer)

cmdSave.Enabled = True

End Sub


Private Sub txtSG_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

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

cmdSave.Enabled = True

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

cmdSave.Enabled = True

End Sub


Private Sub txtUrobilinogen_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub txtWCC_KeyPress(KeyAscii As Integer)

Select Case Chr(KeyAscii)
  Case "P", "p":
    txtWCC = "Packed"
    KeyAscii = 0
  Case ">", "G", "g":
    txtWCC = ">200"
    KeyAscii = 0
  Case "N", "n":
    txtWCC = "Nil"
    KeyAscii = 0
End Select

cmdSave.Enabled = True

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







Private Sub txtWCC_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

pBar = 0

End Sub

Private Sub udSampleID_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

LoadAllDetails

If txtName.Enabled Then
  txtName.SetFocus
End If

pBar = 0

End Sub



Public Property Let FromElseWHERE(ByVal bNewValue As Boolean)

mFromElseWHERE = bNewValue

End Property
Public Property Let FromElseWHERERunNumber(ByVal NewValue As String)

mFromElseWHERERunNumber = NewValue

End Property

Public Property Let UpdatingSensitivity(ByVal vNewValue As Boolean)

mUpdatingSensitivity = vNewValue

End Property
Public Property Let UpdatingExtra(ByVal vNewValue As Boolean)

mUpdatingExtra = vNewValue

End Property

