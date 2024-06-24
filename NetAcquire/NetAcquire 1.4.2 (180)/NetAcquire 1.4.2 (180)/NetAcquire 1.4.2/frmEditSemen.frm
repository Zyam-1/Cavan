VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditSemen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Semen Analysis"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   14040
      MaxLength       =   8
      TabIndex        =   110
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdViewReports 
      BackColor       =   &H00FFFF00&
      Caption         =   "View Reports"
      Height          =   705
      Left            =   12150
      Picture         =   "frmEditSemen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2880
      Width           =   1275
   End
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   765
      Left            =   12150
      Picture         =   "frmEditSemen.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Log as Phoned"
      Top             =   1950
      Width           =   1275
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
      Height          =   420
      Left            =   14040
      TabIndex        =   108
      Top             =   1080
      Width           =   1605
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "&Validate"
      Height          =   735
      Left            =   4680
      Picture         =   "frmEditSemen.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8130
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      Height          =   5265
      Left            =   5910
      TabIndex        =   76
      Top             =   3660
      Width           =   6045
      Begin VB.Frame Frame14 
         Caption         =   "Sperm Morphology"
         Height          =   585
         Left            =   540
         TabIndex        =   105
         Top             =   2940
         Width           =   4935
         Begin VB.TextBox txtMorphology 
            Height          =   285
            Left            =   540
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% Normal"
            Height          =   195
            Index           =   0
            Left            =   1470
            TabIndex        =   106
            Top             =   270
            Width           =   660
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Specimen Type"
         Height          =   585
         Left            =   540
         TabIndex        =   104
         Top             =   0
         Width           =   4935
         Begin VB.ComboBox cmbSpecimenType 
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   210
            Width           =   4695
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "pH"
         Height          =   675
         Left            =   540
         TabIndex        =   103
         Top             =   2220
         Width           =   1845
         Begin VB.TextBox txtpH 
            Height          =   285
            Left            =   150
            TabIndex        =   25
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.ComboBox cmbSemenComment 
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Text            =   "cmbSemenComment"
         Top             =   3810
         Width           =   5415
      End
      Begin VB.Frame Frame8 
         Caption         =   "Comments (Infertility + Post Vasectomy)"
         Height          =   1605
         Left            =   120
         TabIndex        =   93
         Top             =   3570
         Width           =   5715
         Begin VB.TextBox txtSemenComment 
            BackColor       =   &H80000018&
            Height          =   945
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   94
            Top             =   570
            Width           =   5415
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Motility"
         Height          =   1575
         Left            =   2670
         TabIndex        =   82
         Top             =   1320
         Width           =   2805
         Begin VB.TextBox txtMotility 
            Height          =   285
            Index           =   3
            Left            =   690
            TabIndex        =   101
            Top             =   1200
            Width           =   630
         End
         Begin VB.TextBox txtMotility 
            Height          =   285
            Index           =   0
            Left            =   690
            TabIndex        =   85
            Top             =   210
            Width           =   600
         End
         Begin VB.TextBox txtMotility 
            Height          =   285
            Index           =   1
            Left            =   690
            TabIndex        =   84
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtMotility 
            Height          =   285
            Index           =   2
            Left            =   690
            TabIndex        =   83
            Top             =   870
            Width           =   600
         End
         Begin ComCtl2.UpDown udMotility 
            Height          =   285
            Index           =   2
            Left            =   1290
            TabIndex        =   86
            Top             =   870
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   327681
            BuddyControl    =   "txtMotility(2)"
            BuddyDispid     =   196625
            BuddyIndex      =   2
            OrigLeft        =   810
            OrigTop         =   1140
            OrigRight       =   1050
            OrigBottom      =   1485
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown udMotility 
            Height          =   285
            Index           =   1
            Left            =   1290
            TabIndex        =   87
            Top             =   540
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   327681
            BuddyControl    =   "txtMotility(1)"
            BuddyDispid     =   196625
            BuddyIndex      =   1
            OrigLeft        =   810
            OrigTop         =   780
            OrigRight       =   1050
            OrigBottom      =   1065
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown udMotility 
            Height          =   345
            Index           =   0
            Left            =   1290
            TabIndex        =   88
            Top             =   180
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   609
            _Version        =   327681
            BuddyControl    =   "txtMotility(0)"
            BuddyDispid     =   196625
            BuddyIndex      =   0
            OrigLeft        =   810
            OrigTop         =   270
            OrigRight       =   1050
            OrigBottom      =   555
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown udMotility 
            Height          =   285
            Index           =   3
            Left            =   1290
            TabIndex        =   102
            Top             =   1200
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   327681
            BuddyControl    =   "txtMotility(3)"
            BuddyDispid     =   196625
            BuddyIndex      =   3
            OrigLeft        =   810
            OrigTop         =   1140
            OrigRight       =   1050
            OrigBottom      =   1485
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "% Grade A"
            Height          =   195
            Index           =   0
            Left            =   1560
            TabIndex        =   92
            Top             =   270
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "% Grade B"
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   91
            Top             =   600
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "% Grade C"
            Height          =   195
            Index           =   4
            Left            =   1560
            TabIndex        =   90
            Top             =   930
            Width           =   750
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "% Grade D"
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   89
            Top             =   1230
            Width           =   765
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Volume"
         Height          =   675
         Left            =   540
         TabIndex        =   80
         Top             =   1320
         Width           =   1845
         Begin VB.ComboBox cmbVolume 
            Height          =   315
            Left            =   150
            TabIndex        =   24
            Text            =   "cmbVolume"
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "mL"
            Height          =   195
            Index           =   1
            Left            =   1350
            TabIndex        =   81
            Top             =   300
            Width           =   210
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Sperm Count"
         Height          =   675
         Index           =   0
         Left            =   2670
         TabIndex        =   78
         Top             =   600
         Width           =   2805
         Begin VB.ComboBox cmbCount 
            Height          =   315
            Left            =   90
            TabIndex        =   23
            Text            =   "cmbCount"
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Million per mL"
            Height          =   195
            Index           =   0
            Left            =   1380
            TabIndex        =   79
            Top             =   300
            Width           =   960
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Viscosity"
         Height          =   675
         Left            =   540
         TabIndex        =   77
         Top             =   600
         Width           =   1845
         Begin VB.ComboBox cmbConsistency 
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Text            =   "cmbConsistency"
            Top             =   240
            Width           =   1545
         End
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Date"
      Height          =   1785
      Left            =   5910
      TabIndex        =   73
      Top             =   1860
      Width           =   6045
      Begin VB.Frame Frame5 
         Height          =   1125
         Left            =   4680
         TabIndex        =   95
         Top             =   0
         Width           =   1365
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Out of Hours"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   97
            Top             =   690
            Width           =   1215
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Routine"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   96
            Top             =   360
            Width           =   885
         End
      End
      Begin MSComCtl2.DTPicker dtRunDate 
         Height          =   315
         Left            =   540
         TabIndex        =   7
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219152385
         CurrentDate     =   36942
      End
      Begin MSComCtl2.DTPicker dtSampleDate 
         Height          =   315
         Left            =   2370
         TabIndex        =   8
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219152385
         CurrentDate     =   36942
      End
      Begin MSMask.MaskEdBox tSampleTime 
         Height          =   315
         Left            =   3750
         TabIndex        =   9
         ToolTipText     =   "Time of Sample"
         Top             =   360
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
         Left            =   2370
         TabIndex        =   10
         Top             =   1080
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219217921
         CurrentDate     =   38078
      End
      Begin MSMask.MaskEdBox tRecTime 
         Height          =   315
         Left            =   3750
         TabIndex        =   11
         ToolTipText     =   "Time of Sample"
         Top             =   1080
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
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Received in Lab"
         Height          =   195
         Index           =   16
         Left            =   540
         TabIndex        =   107
         Top             =   1140
         Width           =   1770
      End
      Begin VB.Image iToday 
         Height          =   330
         Index           =   2
         Left            =   2820
         Picture         =   "frmEditSemen.frx":114E
         Stretch         =   -1  'True
         ToolTipText     =   "Set to Today"
         Top             =   1410
         Width           =   360
      End
      Begin VB.Image iRecDate 
         Height          =   330
         Index           =   0
         Left            =   2340
         Picture         =   "frmEditSemen.frx":1590
         Stretch         =   -1  'True
         ToolTipText     =   "Previous Day"
         Top             =   1410
         Width           =   480
      End
      Begin VB.Image iRecDate 
         Height          =   330
         Index           =   1
         Left            =   3210
         Picture         =   "frmEditSemen.frx":19D2
         Stretch         =   -1  'True
         ToolTipText     =   "Next Day"
         Top             =   1410
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sample"
         Height          =   195
         Index           =   2
         Left            =   2370
         TabIndex        =   75
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Run"
         Height          =   195
         Index           =   2
         Left            =   540
         TabIndex        =   74
         Top             =   180
         Width           =   300
      End
      Begin VB.Image iRunDate 
         Height          =   330
         Index           =   0
         Left            =   510
         Picture         =   "frmEditSemen.frx":1E14
         Stretch         =   -1  'True
         ToolTipText     =   "Previous Day"
         Top             =   690
         Width           =   480
      End
      Begin VB.Image iRunDate 
         Height          =   330
         Index           =   1
         Left            =   1410
         Picture         =   "frmEditSemen.frx":2256
         Stretch         =   -1  'True
         ToolTipText     =   "Next Day"
         Top             =   690
         Width           =   480
      End
      Begin VB.Image iSampleDate 
         Height          =   330
         Index           =   0
         Left            =   2370
         Picture         =   "frmEditSemen.frx":2698
         Stretch         =   -1  'True
         ToolTipText     =   "Previous Day"
         Top             =   690
         Width           =   480
      End
      Begin VB.Image iSampleDate 
         Height          =   330
         Index           =   1
         Left            =   3240
         Picture         =   "frmEditSemen.frx":2ADA
         Stretch         =   -1  'True
         ToolTipText     =   "Next Day"
         Top             =   690
         Width           =   480
      End
      Begin VB.Image iToday 
         Height          =   330
         Index           =   0
         Left            =   1020
         Picture         =   "frmEditSemen.frx":2F1C
         Stretch         =   -1  'True
         ToolTipText     =   "Set to Today"
         Top             =   690
         Width           =   360
      End
      Begin VB.Image iToday 
         Height          =   330
         Index           =   1
         Left            =   2850
         Picture         =   "frmEditSemen.frx":335E
         Stretch         =   -1  'True
         ToolTipText     =   "Set to Today"
         Top             =   690
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdSaveHold 
      Caption         =   "Save && &Hold"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1830
      Picture         =   "frmEditSemen.frx":37A0
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8130
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   735
      Left            =   3240
      Picture         =   "frmEditSemen.frx":3E0A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8130
      Width           =   1155
   End
   Begin VB.CommandButton cmdOrderSemen 
      Caption         =   "Order Analysis"
      Height          =   735
      Left            =   420
      Picture         =   "frmEditSemen.frx":4474
      Style           =   1  'Graphical
      TabIndex        =   72
      Tag             =   "bOrder"
      Top             =   8130
      Width           =   1155
   End
   Begin VB.Frame Frame4 
      Height          =   6075
      Left            =   390
      TabIndex        =   55
      Top             =   1860
      Width           =   5445
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
         TabIndex        =   109
         ToolTipText     =   "Copy To"
         Top             =   2910
         Width           =   375
      End
      Begin VB.ComboBox cmbHospital 
         Height          =   315
         Left            =   1050
         TabIndex        =   14
         Text            =   "cmbHospital"
         Top             =   2520
         Width           =   3915
      End
      Begin VB.ComboBox cmbDemogComment 
         Height          =   315
         Left            =   1050
         TabIndex        =   18
         Text            =   "cmbDemogComment"
         Top             =   4080
         Width           =   3915
      End
      Begin VB.TextBox txtDemographicComment 
         Height          =   915
         Left            =   1050
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   4440
         Width           =   3885
      End
      Begin VB.ComboBox cmbClinDetails 
         Height          =   315
         Left            =   1050
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   5430
         Width           =   3915
      End
      Begin VB.ComboBox cmbWard 
         Height          =   315
         Left            =   1050
         TabIndex        =   15
         Text            =   "cmbWard"
         Top             =   2910
         Width           =   3915
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   0
         Left            =   750
         MaxLength       =   30
         TabIndex        =   12
         Top             =   1830
         Width           =   4215
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   1
         Left            =   750
         MaxLength       =   30
         TabIndex        =   13
         Top             =   2100
         Width           =   4215
      End
      Begin VB.ComboBox cmbClinician 
         Height          =   315
         Left            =   1020
         TabIndex        =   16
         Text            =   "cmbClinician"
         Top             =   3300
         Width           =   3915
      End
      Begin VB.ComboBox cmbGP 
         Height          =   315
         Left            =   1050
         TabIndex        =   17
         Text            =   "cmbGP"
         Top             =   3690
         Width           =   3915
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Hospital"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   98
         Top             =   2580
         Width           =   570
      End
      Begin VB.Label lblSex 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4260
         TabIndex        =   71
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label lblAge 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3180
         TabIndex        =   70
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label lblDoB 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   69
         Top             =   1170
         Width           =   1515
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   750
         TabIndex        =   68
         Top             =   780
         Width           =   4215
      End
      Begin VB.Label lblChart 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   67
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cl Details"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   66
         Top             =   5490
         Width           =   660
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Chart #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   65
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   64
         Top             =   810
         Width           =   420
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "D.o.B"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   63
         Top             =   1230
         Width           =   405
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Age"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   62
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sex"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3930
         TabIndex        =   61
         Top             =   1200
         Width           =   270
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ward"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   60
         Top             =   2970
         Width           =   390
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Address"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   780
         TabIndex        =   59
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comments"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   255
         TabIndex        =   58
         Top             =   4140
         Width           =   735
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clinician"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   405
         TabIndex        =   57
         Top             =   3360
         Width           =   585
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "GP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   765
         TabIndex        =   56
         Top             =   3750
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdSemenHistory 
      Caption         =   "&History"
      Height          =   735
      Left            =   12150
      Picture         =   "frmEditSemen.frx":477E
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7200
      Width           =   1275
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   615
      Left            =   12150
      Picture         =   "frmEditSemen.frx":4BC0
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   3630
      Width           =   1275
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   11520
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   390
      TabIndex        =   49
      Top             =   120
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
      Height          =   735
      Left            =   12150
      Picture         =   "frmEditSemen.frx":4ECA
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4500
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1395
      Left            =   390
      TabIndex        =   40
      Top             =   330
      Width           =   11565
      Begin VB.TextBox txtForeName 
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
         Left            =   6480
         TabIndex        =   3
         Top             =   570
         Width           =   2115
      End
      Begin VB.TextBox txtSurName 
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
         Left            =   4290
         TabIndex        =   2
         Top             =   570
         Width           =   2175
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
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   4
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   9240
         MaxLength       =   4
         TabIndex        =   5
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   9240
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "Male"
         Top             =   990
         Width           =   1545
      End
      Begin VB.Frame Frame6 
         Height          =   1395
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   2715
         Begin VB.ComboBox cMRU 
            Height          =   315
            Left            =   570
            TabIndex        =   50
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
            TabIndex        =   44
            Top             =   510
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   847
            _Version        =   327681
            Value           =   1
            BuddyControl    =   "txtSampleID"
            BuddyDispid     =   196694
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
            TabIndex        =   51
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image iRelevant 
            Height          =   480
            Index           =   1
            Left            =   1500
            Picture         =   "frmEditSemen.frx":5534
            Top             =   120
            Width           =   480
         End
         Begin VB.Image iRelevant 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   120
            Picture         =   "frmEditSemen.frx":583E
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Sample ID"
            Height          =   195
            Left            =   720
            TabIndex        =   45
            Top             =   30
            Width           =   735
         End
      End
      Begin VB.CommandButton bsearch 
         Appearance      =   0  'Flat
         Caption         =   "Searc&h"
         Height          =   345
         Left            =   7920
         TabIndex        =   42
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton bDoB 
         Appearance      =   0  'Flat
         Caption         =   "Search"
         Height          =   285
         Left            =   10800
         TabIndex        =   41
         Top             =   270
         Width           =   705
      End
      Begin VB.Label lblSurNameTitle 
         AutoSize        =   -1  'True
         Caption         =   "SurName"
         Height          =   195
         Left            =   4290
         TabIndex        =   100
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblForeNameTitle 
         AutoSize        =   -1  'True
         Caption         =   "ForeName"
         Height          =   195
         Left            =   6840
         TabIndex        =   99
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblChartNumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Chart #"
         Height          =   285
         Left            =   2730
         TabIndex        =   54
         ToolTipText     =   "Click to change Location"
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label lAddWardGP 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2730
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   210
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   0
         Left            =   8790
         TabIndex        =   48
         Top             =   300
         Width           =   405
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   0
         Left            =   8880
         TabIndex        =   47
         Top             =   660
         Width           =   285
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Left            =   8910
         TabIndex        =   46
         Top             =   1020
         Width           =   270
      End
   End
   Begin VB.CommandButton bViewBB 
      Caption         =   "Transfusion Details"
      Height          =   885
      Left            =   12150
      Picture         =   "frmEditSemen.frx":5B48
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   420
      Width           =   1275
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   735
      Left            =   12150
      Picture         =   "frmEditSemen.frx":5E52
      Style           =   1  'Graphical
      TabIndex        =   36
      Tag             =   "bprint"
      Top             =   5250
      Width           =   1275
   End
   Begin VB.CommandButton bFAX 
      Caption         =   "FAX"
      Height          =   765
      Index           =   0
      Left            =   12150
      Picture         =   "frmEditSemen.frx":64BC
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6000
      Width           =   1275
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   12150
      Picture         =   "frmEditSemen.frx":68FE
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8220
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditSemen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mNewRecord As Boolean

Private Activated As Boolean

Private pPrintToPrinter As String

Private SampleIDWithOffset As Long
Private FormLoaded As Boolean

Private Sub cmdCopyTo_Click()

          Dim s As String

20500     s = cmbWard & " " & cmbClinician
20510     s = Trim$(s) & " " & cmbGP
20520     s = Trim$(s)

20530     frmCopyTo.lblOriginal = s
20540     frmCopyTo.lblSampleID = txtSampleID
20550     frmCopyTo.Dept = "Z"
20560     frmCopyTo.Show 1

20570     CheckCC

End Sub


Private Sub CheckCC()

          Dim sql As String
          Dim tb As Recordset

20580     On Error GoTo CheckCC_Error

20590     cmdCopyTo.Caption = "cc"
20600     cmdCopyTo.Font.Bold = False
20610     cmdCopyTo.BackColor = &H8000000F

20620     If Trim$(txtSampleID) = "" Then Exit Sub

20630     sql = "Select * from SendCopyTo where " & _
              "SampleID = '" & sysOptSemenOffset(0) + Val(txtSampleID) & "'"
20640     Set tb = New Recordset
20650     RecOpenServer 0, tb, sql
20660     If Not tb.EOF Then
20670         cmdCopyTo.Caption = "++ cc ++"
20680         cmdCopyTo.Font.Bold = True
20690         cmdCopyTo.BackColor = &H8080FF
20700     End If

20710     Exit Sub

CheckCC_Error:

          Dim strES As String
          Dim intEL As Integer

20720     intEL = Erl
20730     strES = Err.Description
20740     LogError "frmEditSemen", "CheckCC", intEL, strES, sql

End Sub

Private Sub cmdPhone_Click()

20750     With frmPhoneLog
20760         .SampleID = Val(txtSampleID) + sysOptSemenOffset(0)
20770         If cmbGP <> "" Then
20780             .GP = cmbGP
20790             .WardOrGP = "GP"
20800         Else
20810             .GP = cmbWard
20820             .WardOrGP = "Ward"
20830         End If
20840         .Show 1
20850     End With

20860     CheckIfPhoned

End Sub


Private Sub cmdViewReports_Click()

          Dim f As Form

20870     Set f = New frmReportViewer

20880     f.Dept = "Semen"
20890     f.SampleID = txtSampleID
20900     f.Show 1

20910     Set f = Nothing

End Sub

Private Sub dtRecDate_CloseUp()

20920     pBar = 0

20930     cmdSave.Enabled = True
20940     cmdSaveHold.Enabled = True

End Sub




Private Sub iRecDate_Click(Index As Integer)

20950     If Index = 0 Then
20960         dtRecDate = DateAdd("d", -1, dtRecDate)
20970     Else
20980         If DateDiff("d", dtRecDate, Now) > 0 Then
20990             dtRecDate = DateAdd("d", 1, dtRecDate)
21000         End If
21010     End If

21020     cmdSave.Enabled = True
21030     cmdSaveHold.Enabled = True

End Sub


Private Sub iToday_Click(Index As Integer)

21040     If Index = 0 Then
21050         dtRunDate = Format$(Now, "dd/mm/yyyy")
21060     ElseIf Index = 1 Then
21070         If DateDiff("d", dtRunDate, Now) > 0 Then
21080             dtSampleDate = dtRunDate
21090         Else
21100             dtSampleDate = Format$(Now, "dd/mm/yyyy")
21110         End If
21120     Else
21130         dtRecDate = Format$(Now, "dd/mm/yyyy")
21140     End If

21150     cmdSave.Enabled = True
21160     cmdSaveHold.Enabled = True

End Sub


Private Sub tRecTime_GotFocus()

21170     tRecTime.SelStart = 0
21180     tRecTime.SelLength = 0

End Sub


Private Sub tRecTime_KeyPress(KeyAscii As Integer)

21190     pBar = 0

21200     cmdSave.Enabled = True
21210     cmdSaveHold.Enabled = True

End Sub



Private Sub CheckPrevious()

          Dim sql As String
          Dim tb As Recordset

21220     On Error GoTo CheckPrevious_Error

21230     If Trim$(txtChart) <> "" Then
              '+++ Junaid 20-05-2024
              '30        sql = "select D.SampleID from Demographics as D, SemenResults50 as S where " & _
              '                "Chart = '" & txtChart & "' and " & _
              '                "D.SampleID < '" & SampleIDWithOffset & "' " & _
              '                "and D.SampleID = S.SampleID"
21240         sql = "select D.SampleID from Demographics as D, SemenResults50 as S where " & _
                  "Chart = '" & txtChart & "' and " & _
                  "D.SampleID < '" & Trim(txtSampleID.Text) & "' " & _
                  "and D.SampleID = S.SampleID"
              '--- Junaid
21250         Set tb = New Recordset
21260         RecOpenServer 0, tb, sql
21270         If Not tb.EOF Then
21280             cmdSemenHistory.Visible = True
21290         Else
21300             cmdSemenHistory.Visible = False
21310         End If
21320     Else
21330         cmdSemenHistory.Visible = False
21340     End If

21350     Exit Sub

CheckPrevious_Error:

          Dim strES As String
          Dim intEL As Integer

21360     intEL = Erl
21370     strES = Err.Description
21380     LogError "frmEditSemen", "CheckPrevious", intEL, strES, sql

End Sub

Private Sub GetSampleIDWithOffset()
          '+++ Junaid 20-05-2024
          '10    SampleIDWithOffset = Val(txtSampleID) + sysOptSemenOffset(0)
21390     SampleIDWithOffset = Val(txtSampleID.Text)
          '--- Junaid

End Sub

Private Sub FillLists()

          Dim sngVol As Single
          Dim Lx As List
          Dim Lxs As New Lists

21400     On Error GoTo FillLists_Error

21410     FillWards cmbWard, HospName(0)
21420     FillClinicians cmbClinician, HospName(0)
21430     FillGPs cmbGP, HospName(0)

21440     cmbSpecimenType.AddItem ""
21450     cmbSpecimenType.AddItem "Infertility Analysis"
21460     cmbSpecimenType.AddItem "Post Vasectomy"

21470     With cmbConsistency
21480         .Clear
21490         .AddItem ""
21500         .AddItem "Low"
21510         .AddItem "Normal"
21520         .AddItem "High"
21530     End With

21540     With cmbVolume
21550         .Clear
21560         .AddItem ""
21570         For sngVol = 10 To 0.5 Step -0.5
21580             .AddItem Format$(sngVol, "0.0")
21590         Next
21600     End With

21610     cmbDemogComment.Clear
21620     Set Lxs = New Lists
21630     Lxs.Load "DE"
21640     For Each Lx In Lxs
21650         cmbDemogComment.AddItem Lx.Text
21660     Next

21670     cmbClinDetails.Clear
21680     Set Lxs = New Lists
21690     Lxs.Load "CD"
21700     For Each Lx In Lxs
21710         cmbClinDetails.AddItem Lx.Text
21720     Next

21730     cmbSemenComment.Clear
21740     Set Lxs = New Lists
21750     Lxs.Load "SE"
21760     For Each Lx In Lxs
21770         cmbSemenComment.AddItem Lx.Text
21780     Next

21790     With cmbCount
21800         .Clear
21810         .AddItem ""
21820         .AddItem "0"
21830         .AddItem "<1"
21840         .AddItem "1-2"
21850         .AddItem "2-5"
21860         .AddItem "5-10"
21870         .AddItem "10-20"
21880         .AddItem "20-50"
21890         .AddItem "50-100"
21900     End With

21910     Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

21920     intEL = Erl
21930     strES = Err.Description
21940     LogError "frmEditSemen", "FillLists", intEL, strES

End Sub


Private Sub LoadComments()

          Dim OB As Observation
          Dim OBs As Observations

21950     On Error GoTo LoadComments_Error

21960     txtSemenComment = ""
21970     txtDemographicComment = ""

21980     If Val(txtSampleID) = 0 Then Exit Sub

21990     Set OBs = New Observations
          '+++ Junaid 20-05-2024
          '60    Set OBs = OBs.Load(SampleIDWithOffset, "Semen", "Demographic")
22000     Set OBs = OBs.Load(Trim(txtSampleID.Text), "Semen", "Demographic")
          '--- Junaid
22010     If Not OBs Is Nothing Then
22020         For Each OB In OBs
22030             Select Case UCase$(OB.Discipline)
                      Case "SEMEN": txtSemenComment = Split_Comm(OB.Comment)
22040                 Case "DEMOGRAPHIC": txtDemographicComment = Split_Comm(OB.Comment)
22050             End Select
22060         Next
22070     End If

22080     Exit Sub

LoadComments_Error:

          Dim strES As String
          Dim intEL As Integer

22090     intEL = Erl
22100     strES = Err.Description
22110     LogError "frmEditSemen", "LoadComments", intEL, strES

End Sub

Private Sub SaveComments()

          Dim OBs As Observations

22120     On Error GoTo SaveComments_Error

22130     txtSampleID = Format(Val(txtSampleID))
22140     If Val(txtSampleID) = 0 Then Exit Sub
22150     Set OBs = New Observations
          '+++ Junaid 20-05-2024
          '50    OBs.Save SampleIDWithOffset, True, _
          '               "Demographic", Trim$(txtDemographicComment), _
          '               "Semen", Trim$(txtSemenComment)
22160     OBs.Save Trim(txtSampleID.Text), True, _
              "Demographic", Trim$(txtDemographicComment), _
              "Semen", Trim$(txtSemenComment)
          '--- Junaid
22170     Exit Sub

SaveComments_Error:

          Dim strES As String
          Dim intEL As Integer

22180     intEL = Erl
22190     strES = Err.Description
22200     LogError "frmEditSemen", "SaveComments", intEL, strES

End Sub

Private Sub bDoB_Click()

22210     pBar = 0

22220     With frmPatHistory
22230         If HospName(0) = "Monaghan" Then
22240             .optBoth = True
22250         Else
22260             .optHistoric = True
22270         End If
22280         .oFor(2) = True
22290         .txtName = txtDoB
22300         .FromEdit = True
22310         .EditScreen = Me
22320         .bsearch = True
          
22330         If Not .NoPreviousDetails Then
22340             .Show 1
22350         Else
22360             FlashNoPrevious Me
22370         End If
22380     End With

End Sub

Private Sub bFAX_Click(Index As Integer)

22390     pBar = 0
22400     If UserHasAuthority(UserMemberOf, "AndrologyFax") = False Then
22410         iMsg "You do not have authority to fax" & vbCrLf & "Please contact system administrator"
22420         Exit Sub
22430     End If

End Sub

Private Sub cmbConsistency_Click()

22440     cmdSave.Enabled = True
22450     cmdSaveHold.Enabled = True

End Sub

Private Sub cmbConsistency_KeyPress(KeyAscii As Integer)

22460     Select Case Chr$(KeyAscii)
              Case "W", "w": cmbConsistency = "Watery"
22470         Case "M", "m": cmbConsistency = "Mucoid"
22480         Case "P", "p": cmbConsistency = "Purulent"
22490         Case Else: cmbConsistency = ""
22500     End Select

22510     KeyAscii = 0

22520     cmdSave.Enabled = True
22530     cmdSaveHold.Enabled = True

End Sub

Private Sub cmbCount_Click()

22540     cmdSave.Enabled = True
22550     cmdSaveHold.Enabled = True

End Sub

Private Sub cmbDemogComment_Click()

22560     txtDemographicComment = Trim$(txtDemographicComment & " " & cmbDemogComment)
22570     cmbDemogComment = ""

22580     cmdSave.Enabled = True
22590     cmdSaveHold.Enabled = True

End Sub


Private Sub cmbDemogComment_LostFocus()

          Dim tb As Recordset
          Dim sql As String

22600     On Error GoTo cmbDemogComment_LostFocus_Error

22610     sql = "Select * from Lists where " & _
              "ListType = 'DE' " & _
              "and Code = '" & cmbDemogComment & "' and InUse = 1"
22620     Set tb = New Recordset
22630     RecOpenServer 0, tb, sql
22640     If Not tb.EOF Then
22650         txtDemographicComment = Trim$(txtDemographicComment & " " & tb!Text & "")
22660     Else
22670         txtDemographicComment = Trim$(txtDemographicComment & " " & cmbDemogComment)
22680     End If
22690     cmbDemogComment = ""

22700     Exit Sub

cmbDemogComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

22710     intEL = Erl
22720     strES = Err.Description
22730     LogError "frmEditSemen", "cmbDemogComment_LostFocus", intEL, strES, sql


End Sub


Private Sub cmbHospital_Click()

22740     FillWards cmbWard, cmbHospital
22750     FillClinicians cmbClinician, cmbHospital
22760     FillGPs cmbGP, cmbHospital

22770     cmdSaveHold.Enabled = True
22780     cmdSave.Enabled = True

End Sub


Private Sub cmbSemenComment_Click()

22790     txtSemenComment = txtSemenComment & cmbSemenComment & " "
22800     cmbSemenComment = ""

End Sub


Private Sub cmbSemenComment_LostFocus()

          Dim tb As Recordset
          Dim sql As String

22810     On Error GoTo cmbSemenComment_LostFocus_Error

22820     sql = "Select * from Lists where " & _
              "ListType = 'SE' " & _
              "and Code = '" & cmbSemenComment & "' and InUse = 1"
22830     Set tb = New Recordset
22840     RecOpenServer 0, tb, sql
22850     If Not tb.EOF Then
22860         txtSemenComment = Trim$(txtSemenComment & " " & tb!Text & "")
22870     Else
22880         txtSemenComment = Trim$(txtSemenComment & " " & cmbSemenComment)
22890     End If
22900     cmbSemenComment = ""

22910     Exit Sub

cmbSemenComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

22920     intEL = Erl
22930     strES = Err.Description
22940     LogError "frmEditSemen", "cmbSemenComment_LostFocus", intEL, strES, sql


End Sub


Private Sub cmbSpecimenType_Click()
          Dim n As Integer

22950     cmbVolume = ""
22960     cmbCount = ""
22970     cmbConsistency = ""
22980     For n = 0 To 3
22990         txtMotility(n) = ""
23000     Next
23010     txtMorphology = ""
23020     txtpH = ""
23030     txtSemenComment = ""

23040     cmdSave.Enabled = True
23050     cmdSaveHold.Enabled = True

23060     If cmbSpecimenType = "Infertility Analysis" Then
23070         LockSemenControls True
23080     ElseIf cmbSpecimenType = "Post Vasectomy" Then
23090         LockSemenControls False
23100     Else
23110         LockSemenControls True
23120     End If

End Sub


Private Sub cmbVolume_Click()

23130     cmdSave.Enabled = True
23140     cmdSaveHold.Enabled = True

End Sub


Private Sub cmdSemenHistory_Click()

23150     With frmSemenReport
23160         .lblChart = txtChart
23170         .lblName = Trim$(txtSurName & " " & txtForeName)
23180         .lblDoB = txtDoB
23190         .lblSex = Trim$(Left$(txtSex & " ", 1))
23200         .lblAge = txtAge
23210         .lblAddress = Trim$(txtAddress(0) & " " & txtAddress(1))
23220         .Show 1
23230     End With

End Sub

Private Sub cmdOrderSemen_Click()
    '
    '      Dim tb As Recordset
    '      Dim sql As String
    '
    '10    On Error GoTo cmdOrderSemen_Click_Error
    '
    '20    GetSampleIDWithOffset
    '
    '30    sql = "Select * from SemenResults where " & _
    '            "SampleID = '" & SampleIDWithOffset & "'"
    '40    Set tb = New Recordset
    '50    RecOpenServer 0, tb, sql
    '60    If tb.EOF Then
    '70      tb.AddNew
    '80      tb!SampleID = SampleIDWithOffset
    '90      tb.Update
    '100   End If
    '
    '110   PBar = 0
    '
    '120   Exit Sub
    '
    'cmdOrderSemen_Click_Error:
    '
    'Dim strES As String
    'Dim intEL As Integer
    '
    '130   intEL = Erl
    '140   strES = Err.Description
    '150   LogError "frmEditSemen", "cmdOrderSemen_Click", intEL, strES, sql
    '
    '
End Sub


Private Sub bPrint_Click()

          Dim tb As Recordset
          Dim sql As String

23240     On Error GoTo bPrint_Click_Error

23250     pBar = 0

23260     If UserHasAuthority(UserMemberOf, "AndrologyPrint") = False Then
23270         iMsg "You do not have authority to print" & vbCrLf & "Please contact system administrator"
23280         Exit Sub
23290     End If

23300     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
23310         Exit Sub
23320     End If
23330     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

23340     GetSampleIDWithOffset
23350     SaveDemographics
23360     SaveSemen gVALID, gNOTPRINTED
          '+++ Junaid 20-05-2024
          '140   sql = "Select * from PrintPending where " & _
          '            "Department = 'Z' " & _
          '            "and SampleID = '" & SampleIDWithOffset & "'"
23370     sql = "Select * from PrintPending where " & _
              "Department = 'Z' " & _
              "and SampleID = '" & Trim(txtSampleID.Text) & "'"
          '--- Junaid
23380     Set tb = New Recordset
23390     RecOpenClient 0, tb, sql
23400     If tb.EOF Then
23410         tb.AddNew
23420     End If
23430     tb!SampleID = txtSampleID
23440     tb!Ward = cmbWard
23450     tb!Clinician = cmbClinician
23460     tb!GP = cmbGP
23470     tb!Department = "Z"
23480     tb!Initiator = UserName
23490     tb!UsePrinter = pPrintToPrinter
23500     tb.Update

23510     txtSampleID = Format$(Val(txtSampleID) + 1)
23520     GetSampleIDWithOffset

23530     LoadAllDetails

23540     Exit Sub

bPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

23550     intEL = Erl
23560     strES = Err.Description
23570     LogError "frmEditSemen", "bPrint_Click", intEL, strES, sql


End Sub

Private Sub SaveDemographics()

          Dim sql As String
          Dim tb As Recordset
          Dim Hosp As String
          Dim NewLabNumber As String

23580     On Error GoTo SaveDemographics_Error

23590     NewLabNumber = DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)

23600     SaveComments

23610     If InStr(UCase$(lblChartNumber), "CAVAN") Then
23620         Hosp = "Cavan"
23630     ElseIf InStr(UCase$(lblChartNumber), "MONAGHAN") Then
23640         Hosp = "Monaghan"
23650     Else
23660         Hosp = ""
23670     End If
          '+++ Junaid 20-05-2024
          '110   sql = "Select * from Demographics where " & _
          '            "SampleID = '" & SampleIDWithOffset & "'"
23680     sql = "Select * from Demographics where " & _
              "SampleID = '" & Trim(txtSampleID.Text) & "'"
          '--- Junaid
23690     Set tb = New Recordset
23700     RecOpenClient 0, tb, sql
23710     If tb.EOF Then
23720         tb.AddNew
23730         tb!Fasting = 0
23740         tb!FAXed = 0
23750     End If

23760     tb!RooH = cRooH(0)

23770     tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
23780     If IsDate(tSampleTime) Then
23790         tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
23800     Else
23810         tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
23820     End If
23830     If IsDate(tRecTime) Then
23840         tb!RecDate = Format$(dtRecDate & " " & tRecTime, "dd/mmm/yyyy hh:mm")
23850     Else
23860         tb!RecDate = Format$(dtRecDate, "dd/mmm/yyyy")
23870     End If
          '+++ Junaid 20-05-2024
          '310   tb!SampleID = SampleIDWithOffset
23880     tb!SampleID = Trim(txtSampleID.Text)
          '--- Junaid
23890     tb!Chart = txtChart
23900     tb!PatName = Trim$(txtSurName & " " & txtForeName)
23910     tb!SurName = txtSurName & ""
23920     tb!ForeName = txtForeName & ""
23930     If IsDate(txtDoB) Then
23940         tb!DoB = Format$(txtDoB, "dd/mmm/yyyy")
23950     Else
23960         tb!DoB = Null
23970     End If
23980     tb!LabNo = NewLabNumber
23990     tb!Age = txtAge
24000     tb!Sex = Left$(txtSex, 1)
24010     tb!Addr0 = txtAddress(0)
24020     tb!Addr1 = txtAddress(1)
24030     tb!Ward = Left$(cmbWard, 50)
24040     tb!Clinician = Left$(cmbClinician, 50)
24050     tb!GP = Left$(cmbGP, 50)
24060     tb!ClDetails = Left$(cmbClinDetails, 30)
24070     tb!Hospital = Hosp
24080     tb!RecordDateTime = Format$(Now, "dd/MMM/yyyy HH:nn:ss")
24090     tb.Update
          '+++ Junaid 20-05-2024
          '530   LogTimeOfPrinting SampleIDWithOffset, "D"
24100     LogTimeOfPrinting Trim(txtSampleID.Text), "D"
          '--- Junaid
          '510   ValidateDemo
24110     Screen.MousePointer = 0

24120     Exit Sub

SaveDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

24130     intEL = Erl
24140     strES = Err.Description
24150     LogError "frmEditSemen", "SaveDemographics", intEL, strES, sql

End Sub



Private Sub ValidateDemo()
24160     On Error GoTo ValidateDemo_Error
        
          Dim DVs As New DemogValidations
          Dim DV As New DemogValidation

24170     Set DV = New DemogValidation
24180     DV.SampleID = txtSampleID
24190     DV.EnteredBy = UserName
24200     DV.EnteredDateTime = Format$(Now, "dd/mmm/yyyy hh:mm:ss")
24210     DV.ValidatedBy = UserName
24220     DVs.Add DV
24230     DVs.Save DV

       
24240     Exit Sub

       
ValidateDemo_Error:

          Dim strES As String
          Dim intEL As Integer

24250     intEL = Erl
24260     strES = Err.Description
24270     LogError "frmEditSemen", "ValidateDemo", intEL, strES
End Sub

Private Sub bPrintHold_Click()

          Dim tb As Recordset
          Dim sql As String

24280     On Error GoTo bPrintHold_Click_Error

24290     pBar = 0

24300     If UserHasAuthority(UserMemberOf, "AndrologyPrint") = False Then
24310         iMsg "You do not have authority to print" & vbCrLf & "Please contact system administrator"
24320         Exit Sub
24330     End If

24340     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
24350         Exit Sub
24360     End If
24370     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

24380     GetSampleIDWithOffset
24390     SaveDemographics

24400     SaveSemen gVALID, gNOTPRINTED
          '+++ Junaid 20-05-2024
          '140   sql = "Select * from PrintPending where " & _
          '            "Department = 'Z' " & _
          '            "and SampleID = '" & SampleIDWithOffset & "'"
24410     sql = "Select * from PrintPending where " & _
              "Department = 'Z' " & _
              "and SampleID = '" & Trim(txtSampleID.Text) & "'"
          '--- Junaid
24420     Set tb = New Recordset
24430     RecOpenClient 0, tb, sql
24440     If tb.EOF Then
24450         tb.AddNew
24460     End If
24470     tb!SampleID = txtSampleID
24480     tb!Ward = cmbWard
24490     tb!Clinician = cmbClinician
24500     tb!GP = cmbGP
24510     tb!Department = "Z"
24520     tb!Initiator = UserName
24530     tb!UsePrinter = pPrintToPrinter
24540     tb.Update

24550     Exit Sub

bPrintHold_Click_Error:

          Dim strES As String
          Dim intEL As Integer

24560     intEL = Erl
24570     strES = Err.Description
24580     LogError "frmEditSemen", "bPrintHold_Click", intEL, strES, sql


End Sub

Private Sub cmdSaveHold_Click()

24590     pBar = 0

24600     If UserHasAuthority(UserMemberOf, "AndrologySave") = False Then
24610         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
24620         Exit Sub
24630     End If


          '20    If Not CheckMotilityTotal() Then
          '30      iMsg "Motility sum not 100%", vbExclamation
          '40      Exit Sub
          '50    End If



24640     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
24650         Exit Sub
24660     End If
24670     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

24680     If lblChartNumber.BackColor = vbRed Then
24690         If iMsg("Confirm this Patient has" & vbCrLf & _
                  lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
24700             Exit Sub
24710         End If
24720     End If

24730     cmdSaveHold.Caption = "Saving"

24740     GetSampleIDWithOffset

24750     SaveDemographics
24760     SaveSemen gNOTVALID, gNOTPRINTED
24770     SaveComments
24780     UpdateMRU Me
          'Call LabNoUpdatePrvData(txtChart, Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, Left$(txtSex, 1), txtLabNo)
24790     cmdSaveHold.Caption = "Save && &Hold"
24800     cmdSaveHold.Enabled = False
24810     cmdSave.Enabled = False

End Sub

Private Sub cmdSave_Click()

24820     pBar = 0

24830     If UserHasAuthority(UserMemberOf, "AndrologySave") = False Then
24840         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
24850         Exit Sub
24860     End If

          '20    If Not CheckMotilityTotal() Then
          '30      iMsg "Motility sum not 100%", vbExclamation
          '40      Exit Sub
          '50    End If

24870     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
24880         Exit Sub
24890     End If
24900     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub


24910     If lblChartNumber.BackColor = vbRed Then
24920         If iMsg("Confirm this Patient has" & vbCrLf & _
                  lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
24930             Exit Sub
24940         End If
24950     End If

24960     cmdSaveHold.Caption = "Saving"

24970     GetSampleIDWithOffset

24980     SaveDemographics
24990     SaveSemen gNOTVALID, gNOTPRINTED
25000     SaveComments
25010     UpdateMRU Me
          'Call LabNoUpdatePrvData(txtChart, Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, Left$(txtSex, 1), txtLabNo)
25020     cmdSaveHold.Caption = "Save && &Hold"
25030     cmdSaveHold.Enabled = False
25040     cmdSave.Enabled = False

25050     txtSampleID = Format$(Val(txtSampleID) + 1)
25060     GetSampleIDWithOffset

25070     LoadAllDetails

25080     cmdSave.Enabled = True
25090     cmdSaveHold.Enabled = True

End Sub

Private Sub bsearch_Click()

25100     pBar = 0

25110     With frmPatHistory
25120         If HospName(0) = "Monaghan" Then
25130             .optBoth = True
25140         Else
25150             .optHistoric = True
25160         End If
25170         .oFor(0) = True
25180         .txtName = Trim$(txtSurName & " " & txtForeName)
25190         .FromEdit = True
25200         .EditScreen = Me
25210         .bsearch = True
25220         If Not .NoPreviousDetails Then
25230             .Show 1
25240         Else
25250             FlashNoPrevious Me
25260         End If
25270     End With
25280     LabNoUpdatePrvColor
End Sub

Private Sub cmdValidate_Click()

          Dim tb As Recordset
          Dim sql As String

25290     On Error GoTo cmdValidate_Click_Error

25300     pBar = 0

25310     If UserHasAuthority(UserMemberOf, "AndrologySave") = False Then
25320         iMsg "You do not have authority to save" & vbCrLf & "Please contact system administrator"
25330         Exit Sub
25340     End If

25350     If Not EntriesOK(txtSampleID, txtSurName, txtSex, cmbWard.Text, cmbGP.Text, cmbClinician.Text) Then
25360         Exit Sub
25370     End If
25380     If Not CheckTimes(tSampleTime, txtDemographicComment, tRecTime, cmbHospital, cmbWard) Then Exit Sub

25390     If lblChartNumber.BackColor = vbRed Then
25400         If iMsg("Confirm this Patient has" & vbCrLf & _
                  lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
25410             Exit Sub
25420         End If
25430     End If

25440     GetSampleIDWithOffset

25450     SaveDemographics
25460     SaveComments
25470     UpdateMRU Me

25480     If cmdValidate.Caption = "&Validate" Then
25490         SaveSemen gVALID, gNOTPRINTED
              '+++ Junaid 20-05-2024
              '220       sql = "Select * from PrintPending where " & _
              '                "Department = 'Z' " & _
              '                "and SampleID = '" & SampleIDWithOffset & "'"
25500         sql = "Select * from PrintPending where " & _
                  "Department = 'Z' " & _
                  "and SampleID = '" & Trim(txtSampleID.Text) & "'"
              '--- Junaid
25510         Set tb = New Recordset
25520         RecOpenClient 0, tb, sql
25530         If tb.EOF Then
25540             tb.AddNew
25550         End If
25560         tb!SampleID = txtSampleID
25570         tb!Ward = cmbWard
25580         tb!Clinician = cmbClinician
25590         tb!GP = cmbGP
25600         tb!Department = "Z"
25610         tb!Initiator = UserName
25620         tb!UsePrinter = pPrintToPrinter
25630         tb.Update
25640         txtSampleID = Format$(Val(txtSampleID) + 1)
25650     Else
25660         SaveSemen gNOTVALID, gNOTPRINTED
25670     End If

25680     GetSampleIDWithOffset
25690     LoadAllDetails

25700     Exit Sub

cmdValidate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

25710     intEL = Erl
25720     strES = Err.Description
25730     LogError "frmEditSemen", "cmdValidate_Click", intEL, strES, sql


End Sub

Private Sub bViewBB_Click()

25740     pBar = 0

25750     If Trim$(txtChart) <> "" Then
25760         frmViewBB.lchart = txtChart
25770         frmViewBB.Show 1
25780     End If

End Sub


Private Sub LoadAllDetails()

25790     LoadDemographics
25800     LoadSemen

25810     LoadComments

25820     SetViewReports txtSampleID

25830     If cmbSpecimenType = "Infertility Analysis" Then
25840         LockSemenControls True
25850     ElseIf cmbSpecimenType = "Post Vasectomy" Then
25860         LockSemenControls False
25870     Else
25880         LockSemenControls True
25890     End If

25900     CheckCC
25910     LabNoUpdatePrvColor
End Sub
Private Sub bcancel_Click()

25920     pBar = 0

25930     Unload Me

End Sub

Private Sub cmbClinDetails_Click()

25940     cmdSaveHold.Enabled = True
25950     cmdSave.Enabled = True

End Sub


Private Sub cmbClinDetails_LostFocus()

          Dim tb As Recordset
          Dim sql As String

25960     On Error GoTo cmbClinDetails_LostFocus_Error

25970     pBar = 0

25980     If Trim$(cmbClinDetails) = "" Then Exit Sub

25990     sql = "Select * from Lists where " & _
              "ListType = 'CD' and " & _
              "Code = '" & AddTicks(Trim$(cmbClinDetails)) & "' and InUse = 1"
26000     Set tb = New Recordset
26010     RecOpenServer 0, tb, sql
26020     If Not tb.EOF Then
26030         cmbClinDetails = tb!Text & ""
26040     End If

26050     Exit Sub

cmbClinDetails_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

26060     intEL = Erl
26070     strES = Err.Description
26080     LogError "frmEditSemen", "cmbClinDetails_LostFocus", intEL, strES, sql


End Sub


Private Sub cmbClinician_Click()

26090     cmdSaveHold.Enabled = True
26100     cmdSave.Enabled = True

End Sub

Private Sub cmbClinician_KeyPress(KeyAscii As Integer)

26110     cmdSaveHold.Enabled = True
26120     cmdSave.Enabled = True

End Sub


Private Sub cmbClinician_LostFocus()

26130     pBar = 0
26140     cmbClinician = QueryKnown(cmbClinician, HospName(0))

End Sub

Private Sub cmbGP_Change()

26150     lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

End Sub

Private Sub cmbGP_Click()

26160     pBar = 0

26170     lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

26180     cmdSaveHold.Enabled = True
26190     cmdSave.Enabled = True

End Sub


Private Sub cmbGP_KeyPress(KeyAscii As Integer)

26200     cmdSaveHold.Enabled = True
26210     cmdSave.Enabled = True

End Sub


Private Sub cmbGP_LostFocus()

          Dim strOrig As String
          Dim Gx As New GP

26220     pBar = 0

26230     strOrig = cmbGP

26240     cmbGP = ""

26250     Gx.LoadCodeOrText strOrig
26260     cmbGP = Gx.Text
26270     If sysOptAllowGPFreeText(0) And cmbGP = "" Then
26280         cmbGP = strOrig
26290     End If

End Sub


Private Sub cmdSetPrinter_Click()

26300     frmForcePrinter.From = Me
26310     frmForcePrinter.Show 1

26320     If pPrintToPrinter = "Automatic Selection" Then
26330         pPrintToPrinter = ""
26340     End If

26350     If pPrintToPrinter <> "" Then
26360         cmdSetPrinter.BackColor = vbRed
26370         cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
26380     Else
26390         cmdSetPrinter.BackColor = vbButtonFace
26400         pPrintToPrinter = ""
26410         cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
26420     End If

End Sub

Private Sub cMRU_Click()

26430     txtSampleID = cMRU
26440     GetSampleIDWithOffset

26450     LoadAllDetails

26460     cmdSaveHold.Enabled = False
26470     cmdSave.Enabled = False

End Sub


Private Sub cMRU_KeyPress(KeyAscii As Integer)

26480     KeyAscii = 0

End Sub


Private Sub cRooH_Click(Index As Integer)

26490     cmdSaveHold.Enabled = True
26500     cmdSave.Enabled = True

End Sub

Private Sub cmbWard_Change()

26510     lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

End Sub

Private Sub cmbWard_Click()

26520     lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

26530     cmdSaveHold.Enabled = True
26540     cmdSave.Enabled = True

End Sub


Private Sub cmbWard_KeyPress(KeyAscii As Integer)

26550     cmdSaveHold.Enabled = True
26560     cmdSave.Enabled = True

End Sub


Private Sub cmbWard_LostFocus()

          Dim Hospital As String

26570     Hospital = ListCodeFor("HO", HospName(0))

26580     cmbWard = GetWard(cmbWard, Hospital)

26590     If Trim$(cmbWard) = "" Then
26600         cmbWard = "GP"
26610         Exit Sub
26620     End If

End Sub


Private Sub dtRunDate_CloseUp()

26630     pBar = 0

26640     cmdSaveHold.Enabled = True
26650     cmdSave.Enabled = True

End Sub


Private Sub dtSampleDate_CloseUp()

26660     pBar = 0

26670     cmdSaveHold.Enabled = True
26680     cmdSave.Enabled = True

End Sub


Private Sub Form_Activate()

26690     pBar = 0
26700     pBar.max = LogOffDelaySecs
26710     TimerBar.Enabled = True

End Sub

Private Sub SetViewReports(ByVal SampleID As String)

          Dim sql As String
          Dim tb As New Recordset

26720     On Error GoTo SetViewReports_Error

26730     cmdViewReports.Visible = False

26740     sql = "SELECT COUNT(*) Tot FROM Reports " & _
              "WHERE SampleID = '" & SampleID & "' " & _
              "AND Dept = 'Semen'"
26750     Set tb = Cnxn(0).Execute(sql)
26760     cmdViewReports.Visible = tb!Tot > 0

26770     Exit Sub

SetViewReports_Error:

          Dim strES As String
          Dim intEL As Integer

26780     intEL = Erl
26790     strES = Err.Description
26800     LogError "frmEditSemen", "SetViewReports", intEL, strES, sql

End Sub

Private Sub CheckIfPhoned()

          Dim s As String
          Dim PhLog As PhoneLog
          Dim OBs As Observations

26810     On Error GoTo CheckIfPhoned_Error

26820     PhLog = CheckPhoneLog(Val(txtSampleID) + sysOptSemenOffset(0))
26830     If PhLog.SampleID <> 0 Then
26840         cmdPhone.BackColor = vbYellow
26850         cmdPhone.Caption = "Results Phoned"
26860         cmdPhone.ToolTipText = "Results Phoned"
26870         If InStr(txtDemographicComment.Text, "Results Phoned") = 0 Then
26880             s = "Results Phoned to " & PhLog.PhonedTo & " at " & _
                      Format$(PhLog.DateTime, "hh:mm") & " on " & Format$(PhLog.DateTime, "dd/MM/yyyy") & _
                      " by " & PhLog.PhonedBy & "."
26890             If Trim$(txtDemographicComment.Text) = "" Then
26900                 txtDemographicComment.Text = s
26910             Else
26920                 txtDemographicComment.Text = txtDemographicComment.Text & ". " & s
26930             End If
26940             Set OBs = New Observations
26950             OBs.Save PhLog.SampleID, True, "Demographic", txtDemographicComment.Text

26960         End If
26970     Else
26980         cmdPhone.BackColor = &H8000000F
26990         cmdPhone.Caption = "Phone Results"
27000         cmdPhone.ToolTipText = "Phone Results"
27010     End If

27020     Exit Sub

CheckIfPhoned_Error:

          Dim strES As String
          Dim intEL As Integer

27030     intEL = Erl
27040     strES = Err.Description
27050     LogError "frmEditSemen", "CheckIfPhoned", intEL, strES

End Sub


Private Sub SaveSemen(ByVal Valid As Integer, ByVal Printed As Integer)

          Dim SR As New SemenResult
          '+++ Junaid 20-05-2024
          '10    SR.SampleID = SampleIDWithOffset
27060     SR.SampleID = Trim(txtSampleID.Text)
          '--- Junaid
27070     SR.UserName = UserName
27080     SR.Valid = Valid
27090     If Valid Then
27100         SR.ValidatedBy = UserName
27110         SR.ValidatedDateTime = Now
27120     Else
27130         SR.ValidatedBy = ""
27140         SR.ValidatedDateTime = ""
27150     End If
27160     SR.Printed = Printed
27170     If Printed Then
27180         SR.PrintedBy = UserName
27190         SR.PrintedDateTime = Now
27200     Else
27210         SR.PrintedBy = ""
27220         SR.PrintedDateTime = ""
27230     End If

27240     SR.TestName = "Volume"
27250     SR.Result = Trim$(cmbVolume)
27260     SR.Save

27270     SR.TestName = "Consistency"
27280     SR.Result = Trim$(cmbConsistency)
27290     SR.Save

27300     SR.TestName = "SemenCount"
27310     SR.Result = Trim$(cmbCount)
27320     SR.Save

27330     SR.TestName = "pH"
27340     SR.Result = Trim$(txtpH)
27350     SR.Save

27360     SR.TestName = "GradeA"
27370     SR.Result = Trim$(txtMotility(0))
27380     SR.Save

27390     SR.TestName = "GradeB"
27400     SR.Result = Trim$(txtMotility(1))
27410     SR.Save

27420     SR.TestName = "GradeC"
27430     SR.Result = Trim$(txtMotility(2))
27440     SR.Save

27450     SR.TestName = "GradeD"
27460     SR.Result = Trim$(txtMotility(3))
27470     SR.Save

27480     SR.TestName = "Morphology"
27490     SR.Result = Trim$(txtMorphology)
27500     SR.Save

27510     SR.TestName = "SpecimenType"
27520     SR.Result = Trim$(cmbSpecimenType)
27530     SR.Save

27540     Exit Sub

SaveSemen_Error:

          Dim strES As String
          Dim intEL As Integer

27550     intEL = Erl
27560     strES = Err.Description
27570     LogError "frmEditSemen", "SaveSemen", intEL, strES

End Sub



Private Sub Form_Deactivate()

27580     pBar = 0
27590     TimerBar.Enabled = False

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

27600     pBar = 0

End Sub

Private Sub Form_Load()

27610     FillLists

27620     FillMRU Me

27630     With lblChartNumber
27640         .BackColor = &H8000000F
27650         .ForeColor = vbBlack
27660         Select Case UCase$(HospName(0))
                  Case "CAVAN"
27670                 .Caption = "Cavan Chart #"
27680             Case "MONAGHAN"
27690                 .Caption = "Monaghan Chart #"
27700         End Select
27710     End With

27720     dtRunDate = Format$(Now, "dd/mm/yyyy")
27730     dtSampleDate = Format$(Now, "dd/mm/yyyy")

27740     UpDown1.max = 999999

27750     txtSampleID = GetSetting("NetAcquire", "StartUp", "LastUsedSemen", "1")
27760     GetSampleIDWithOffset
27770     LoadAllDetails

27780     cmdSaveHold.Enabled = False
27790     cmdSave.Enabled = False

27800     Activated = False
27810     FormLoaded = True
End Sub
Private Sub LoadDemographics()

          Dim sql As String
          Dim tb As Recordset
          Dim SampleDate As String
          Dim RooH As Boolean

27820     On Error GoTo LoadDemographics_Error

27830     RooH = IsRoutine()
27840     cRooH(0) = RooH
27850     cRooH(1) = Not RooH
27860     bViewBB.Enabled = False
27870     If FormLoaded Then txtLabNo = ""        'Val(FndMaxID("demographics", "LabNo", ""))
27880     If Trim$(txtSampleID) = "" Then Exit Sub
          '+++ Junaid 20-05-2024
          '80    sql = "Select * from Demographics where " & _
          '            "SampleID = '" & SampleIDWithOffset & "'"
27890     sql = "Select * from Demographics where " & _
              "SampleID = '" & Trim(txtSampleID.Text) & "'"
          '--- Junaid
27900     Set tb = New Recordset
27910     RecOpenClient 0, tb, sql
27920     If tb.EOF Then
27930         mNewRecord = True
27940         dtRunDate = Format$(Now, "dd/mm/yyyy")
27950         dtSampleDate = Format$(Now, "dd/mm/yyyy")
27960         dtRecDate = dtSampleDate
27970         tRecTime.Mask = ""
27980         tRecTime.Text = ""
27990         tRecTime.Mask = "##:##"
28000         txtChart = ""
28010         txtSurName = ""
28020         txtForeName = ""
28030         txtAddress(0) = ""
28040         txtAddress(1) = ""
28050         txtSex = ""
28060         txtDoB = ""
28070         txtAge = ""
28080         cmbWard = "GP"
28090         cmbClinician = ""
28100         cmbGP = ""
28110         cmbHospital = HospName(0)
28120         txtDemographicComment = ""
28130         tSampleTime.Mask = ""
28140         tSampleTime.Text = ""
28150         tSampleTime.Mask = "##:##"
28160         lblChartNumber.Caption = HospName(0) & " Chart #"
28170         lblChartNumber.BackColor = &H8000000F
28180         lblChartNumber.ForeColor = vbBlack
28190         cmbClinDetails = ""
28200     Else
28210         If Trim$(tb!Hospital & "") <> "" Then
28220             cmbHospital = Trim$(tb!Hospital)
28230             lblChartNumber = Trim$(tb!Hospital) & " Chart #"
28240             If UCase$(tb!Hospital) = UCase$(HospName(0)) Then
28250                 lblChartNumber.BackColor = &H8000000F
28260                 lblChartNumber.ForeColor = vbBlack
28270             Else
28280                 lblChartNumber.BackColor = vbRed
28290                 lblChartNumber.ForeColor = vbYellow
28300             End If
28310         Else
28320             cmbHospital = HospName(0)
28330             lblChartNumber.Caption = HospName(0) & " Chart #"
28340             lblChartNumber.BackColor = &H8000000F
28350             lblChartNumber.ForeColor = vbBlack
28360         End If
28370         If IsDate(tb!SampleDate) Then
28380             dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
28390         Else
28400             dtSampleDate = Format$(Now, "dd/mm/yyyy")
28410         End If
28420         If IsDate(tb!Rundate) Then
28430             dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
28440         Else
28450             dtRunDate = Format$(Now, "dd/mm/yyyy")
28460         End If
28470         If IsDate(tb!RecDate) Then
28480             dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
28490         Else
28500             dtRecDate = dtRunDate
28510         End If
28520         mNewRecord = False
28530         If tb!LabNo & "" <> "" Then
28540             txtLabNo = tb!LabNo & ""
28550         End If
28560         cRooH(0) = tb!RooH
28570         cRooH(1) = Not tb!RooH
28580         txtChart = tb!Chart & ""
28590         txtSurName = SurName(tb!PatName & "")
28600         txtForeName = ForeName(tb!PatName & "")
28610         txtAddress(0) = tb!Addr0 & ""
28620         txtAddress(1) = tb!Addr1 & ""
28630         Select Case Left$(Trim$(UCase$(tb!Sex & "")), 1)
                  Case "M": txtSex = "Male"
28640             Case "F": txtSex = "Female"
28650             Case Else: txtSex = ""
28660         End Select
28670         txtDoB = Format$(tb!DoB, "dd/mm/yyyy")
28680         txtAge = tb!Age & ""
28690         cmbWard = tb!Ward & ""
28700         cmbClinician = tb!Clinician & ""
28710         cmbGP = tb!GP & ""
28720         cmbClinDetails = tb!ClDetails & ""
28730         If IsDate(tb!SampleDate) Then
28740             dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
28750             If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
28760                 tSampleTime = Format$(tb!SampleDate, "hh:mm")
28770             Else
28780                 tSampleTime.Mask = ""
28790                 tSampleTime.Text = ""
28800                 tSampleTime.Mask = "##:##"
28810             End If
28820         Else
28830             dtSampleDate = Format$(Now, "dd/mm/yyyy")
28840             tSampleTime.Mask = ""
28850             tSampleTime.Text = ""
28860             tSampleTime.Mask = "##:##"
28870         End If
28880         If IsDate(tb!RecDate & "") Then
28890             dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
28900             If Format$(tb!RecDate, "hh:mm") <> "00:00" Then
28910                 tRecTime = Format$(tb!RecDate, "hh:mm")
28920             Else
28930                 tRecTime.Mask = ""
28940                 tRecTime.Text = ""
28950                 tRecTime.Mask = "##:##"
28960             End If
28970         Else
28980             dtRecDate = dtSampleDate
28990             tRecTime.Mask = ""
29000             tRecTime.Text = ""
29010             tRecTime.Mask = "##:##"
29020         End If
29030     End If

29040     cmdSaveHold.Enabled = False
29050     cmdSave.Enabled = False

29060     If Trim$(txtChart) <> "" Then
29070         sql = "Select  * from PatientDetails where " & _
                  "PatNum = '" & txtChart & "'"
29080         Set tb = New Recordset
29090         RecOpenClientBB 0, tb, sql
29100         bViewBB.Enabled = Not tb.EOF
29110     End If

29120     CheckPrevious

29130     Exit Sub

LoadDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

29140     intEL = Erl
29150     strES = Err.Description
29160     LogError "frmEditSemen", "LoadDemographics", intEL, strES, sql

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29170     pBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

29180     If Val(txtSampleID) > Val(GetSetting("NetAcquire", "StartUp", "LastUsedSemen", "1")) Then
29190         SaveSetting "NetAcquire", "StartUp", "LastUsedSemen", txtSampleID
29200     End If

29210     pPrintToPrinter = ""

29220     Activated = False

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29230     pBar = 0

End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29240     pBar = 0

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29250     pBar = 0

End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29260     pBar = 0

End Sub

Private Sub Frame7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

29270     pBar = 0

End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29280     pBar = 0

End Sub

Private Sub irelevant_Click(Index As Integer)

          Dim sql As String
          Dim tb As Recordset
          Dim strDirection As String

29290     On Error GoTo irelevant_Click_Error

29300     strDirection = IIf(Index = 0, "<", ">")
29310     GetSampleIDWithOffset
          '+++ Junaid 20-05-2024
          '40    sql = "Select top 1 SampleID from SemenResults50 where " & _
          '            "cast(SampleID as numeric) " & strDirection & " '" & SampleIDWithOffset & "' " & _
          '            "Order by SampleID " & IIf(strDirection = "<", "Desc", "Asc")
            
29320     sql = "Select top 1 SampleID from SemenResults50 where " & _
              "cast(SampleID as numeric) " & strDirection & " '" & Trim(txtSampleID.Text) & "' " & _
              "Order by SampleID " & IIf(strDirection = "<", "Desc", "Asc")
          '--- Junaid
29330     Set tb = New Recordset
29340     RecOpenClient 0, tb, sql
29350     If Not tb.EOF Then
              '+++ Junaid 20-05-2024
              '80        txtSampleID = Val(tb!SampleID & "") - sysOptSemenOffset(0)
29360         txtSampleID = Val(tb!SampleID & "")
              '--- Junaid
29370     End If

29380     GetSampleIDWithOffset
29390     LoadAllDetails

29400     cmdSaveHold.Enabled = False
29410     cmdSave.Enabled = False

29420     Exit Sub

irelevant_Click_Error:

          Dim strES As String
          Dim intEL As Integer

29430     intEL = Erl
29440     strES = Err.Description
29450     LogError "frmEditSemen", "irelevant_Click", intEL, strES, sql

End Sub

Private Sub iRunDate_Click(Index As Integer)

29460     If Index = 0 Then
29470         dtRunDate = DateAdd("d", -1, dtRunDate)
29480     Else
29490         If DateDiff("d", dtRunDate, Now) > 0 Then
29500             dtRunDate = DateAdd("d", 1, dtRunDate)
29510         End If
29520     End If

29530     cmdSave.Enabled = True
29540     cmdSaveHold.Enabled = True

End Sub

Private Sub iSampleDate_Click(Index As Integer)

29550     If Index = 0 Then
29560         dtSampleDate = DateAdd("d", -1, dtSampleDate)
29570     Else
29580         If DateDiff("d", dtSampleDate, Now) > 0 Then
29590             dtSampleDate = DateAdd("d", 1, dtSampleDate)
29600         End If
29610     End If

29620     cmdSave.Enabled = True
29630     cmdSaveHold.Enabled = True

End Sub


'Private Sub iToday_Click(index As Integer)
'
'If index = 0 Then
'  dtRunDate = Format$(Now, "dd/mm/yyyy")
'Else
'  If DateDiff("d", dtRunDate, Now) > 0 Then
'    dtSampleDate = dtRunDate
'  Else
'    dtSampleDate = Format$(Now, "dd/mm/yyyy")
'  End If
'End If
'
'cmdSave.Enabled = True
'cmdSaveHold.Enabled = True
'
'End Sub


Private Sub lblChartNumber_Click()
End Sub

'Private Sub lblMotilitySum_Change()
'
'10    If Val(lblMotilitySum) = 0 Then
'20      lblMotilitySum.ForeColor = vbBlack
'30      lblMotilitySum.Font.Bold = False
'40    ElseIf Val(lblMotilitySum) = 100 Then
'50      lblMotilitySum.ForeColor = vbBlack
'60      lblMotilitySum.Font.Bold = False
'70    ElseIf Val(lblMotilitySum) < 100 Then
'80      lblMotilitySum.ForeColor = vbBlue
'90      lblMotilitySum.Font.Bold = True
'100   ElseIf Val(lblMotilitySum) > 100 Then
'110     lblMotilitySum.ForeColor = vbRed
'120     lblMotilitySum.Font.Bold = True
'130   End If
'
'End Sub
'
Private Sub txtaddress_Change(Index As Integer)

29640     lAddWardGP = Trim$(txtAddress(0)) & " : " & cmbWard & " : " & cmbGP

End Sub

Private Sub txtaddress_KeyPress(Index As Integer, KeyAscii As Integer)

29650     cmdSaveHold.Enabled = True
29660     cmdSave.Enabled = True

End Sub


Private Sub txtaddress_LostFocus(Index As Integer)

29670     txtAddress(Index) = Initial2Upper(txtAddress(Index))

End Sub


Private Sub txtage_Change()

29680     lblAge = txtAge

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

29690     cmdSaveHold.Enabled = True
29700     cmdSave.Enabled = True

End Sub


Private Sub txtchart_Change()

29710     lblChart = txtChart

End Sub

Private Sub txtChart_KeyPress(KeyAscii As Integer)

29720     cmdSaveHold.Enabled = True
29730     cmdSave.Enabled = True

End Sub


Private Sub txtchart_LostFocus()

29740     If Trim$(txtChart) = "" Then Exit Sub
29750     If Trim$(txtSurName) <> "" Then Exit Sub

29760     LoadPatientFromChart Me, mNewRecord

End Sub

Private Sub txtMorphology_KeyPress(KeyAscii As Integer)

29770     cmdSave.Enabled = True
29780     cmdSaveHold.Enabled = True

End Sub


Private Sub txtMotility_Change(Index As Integer)
    '
    '10    If Trim(txtMotility(0)) <> "" And _
    '         Trim(txtMotility(1)) <> "" And _
    '         Trim(txtMotility(2)) <> "" Then 'All Filled
    '
    '20      lblMotilitySum = Format(Val(txtMotility(0)) + Val(txtMotility(1)) + Val(txtMotility(2)))
    '
    '30    Else
    '40      lblMotilitySum = ""
    '50    End If
    '
End Sub




Private Sub ClearSemen()

          Dim n As Integer

29790     cmbVolume = ""
29800     cmbCount = ""
29810     cmbConsistency = ""
29820     For n = 0 To 3
29830         txtMotility(n) = ""
29840     Next
29850     cmbSpecimenType = ""
29860     txtMorphology = ""
29870     txtpH = ""
29880     txtSemenComment = ""

End Sub


Private Sub LoadSemen()

          Dim SRS As New SemenResults
          Dim SR As SemenResult

29890     On Error GoTo LoadSemen_Error

29900     ClearSemen

29910     cmdValidate.Caption = "&Validate"
29920     cmdValidate.BackColor = vbButtonFace
          '+++ Junaid 20-05-2024
          '50    SRS.Load SampleIDWithOffset
29930     SRS.Load Trim(txtSampleID.Text)
          '--- JUnaid
29940     For Each SR In SRS
29950         If SR.Valid Then
29960             cmdValidate.Caption = "&Validated"
29970             cmdValidate.BackColor = vbGreen
29980         End If
29990         Select Case UCase$(SR.TestName)
                  Case "VOLUME": cmbVolume = SR.Result
30000             Case "SEMENCOUNT": cmbCount = SR.Result
30010             Case "CONSISTENCY": cmbConsistency = SR.Result
30020             Case "PH": txtpH = SR.Result
30030             Case "GRADEA": txtMotility(0) = SR.Result
30040             Case "GRADEB": txtMotility(1) = SR.Result
30050             Case "GRADEC": txtMotility(2) = SR.Result
30060             Case "GRADED": txtMotility(3) = SR.Result
30070             Case "MORPHOLOGY": txtMorphology = SR.Result
30080             Case "SPECIMENTYPE": cmbSpecimenType = SR.Result
30090         End Select
30100     Next

30110     Exit Sub

LoadSemen_Error:

          Dim strES As String
          Dim intEL As Integer

30120     intEL = Erl
30130     strES = Err.Description
30140     LogError "frmEditSemen", "LoadSemen", intEL, strES

End Sub
Private Sub txtDoB_Change()

30150     lblDoB = txtDoB

30160     LabNoUpdatePrviousData = ""
30170     LabNoUpdatePrvColor

End Sub

Private Sub txtDoB_KeyPress(KeyAscii As Integer)

30180     cmdSaveHold.Enabled = True
30190     cmdSave.Enabled = True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtDoB_LostFocus
' Author    : Masood
' Date      : 09/Jun/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtDoB_LostFocus()

30200     On Error GoTo txtDoB_LostFocus_Error


30210     txtDoB = Convert62Date(txtDoB, BACKWARD)
30220     txtAge = CalcAge(txtDoB, dtSampleDate)
          'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)


30230     Exit Sub


txtDoB_LostFocus_Error:

          Dim strES      As String
          Dim intEL      As Integer

30240     intEL = Erl
30250     strES = Err.Description
30260     LogError "frmEditSemen", "txtDoB_LostFocus", intEL, strES

End Sub


Private Sub TimerBar_Timer()

30270     pBar = pBar + 1

30280     If pBar = pBar.max Then
30290         Unload Me
30300         Exit Sub
30310     End If

End Sub


Private Sub txtpH_KeyPress(KeyAscii As Integer)

30320     cmdSave.Enabled = True
30330     cmdSaveHold.Enabled = True

End Sub


Private Sub txtSex_LostFocus()

    'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)

End Sub

Private Sub txtSurName_Change()

30340     lblName = Trim$(txtSurName & " " & txtForeName)
30350     LabNoUpdatePrviousData = ""
30360     LabNoUpdatePrvColor

End Sub

Private Sub txtForeName_Change()

30370     lblName = Trim$(txtSurName & " " & txtForeName)
30380     LabNoUpdatePrviousData = ""
30390     LabNoUpdatePrvColor
End Sub

Private Sub txtSurName_KeyPress(KeyAscii As Integer)

30400     cmdSaveHold.Enabled = True
30410     cmdSave.Enabled = True

End Sub


Private Sub txtForeName_KeyPress(KeyAscii As Integer)

30420     cmdSaveHold.Enabled = True
30430     cmdSave.Enabled = True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtSurname_LostFocus
' Author    : Masood
' Date      : 09/Jun/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSurname_LostFocus()

          Dim strSurName As String
          Dim strForeName As String
          Dim strSex As String

30440     On Error GoTo txtSurname_LostFocus_Error


30450     strSurName = txtSurName
30460     strForeName = txtForeName
30470     strSex = txtSex

30480     NameLostFocus strSurName, strForeName, strSex

30490     txtSurName = strSurName
30500     txtForeName = strForeName

30510     txtSex = strSex
          'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)

       
30520     Exit Sub

       
txtSurname_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

30530     intEL = Erl
30540     strES = Err.Description
30550     LogError "frmEditSemen", "txtSurname_LostFocus", intEL, strES

End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtForeName_LostFocus
' Author    : Masood
' Date      : 09/Jun/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtForeName_LostFocus()

          Dim strSurName As String
          Dim strForeName As String
          Dim strSex As String

30560     On Error GoTo txtForeName_LostFocus_Error


30570     strSurName = txtSurName
30580     strForeName = txtForeName
30590     strSex = txtSex

30600     NameLostFocus strSurName, strForeName, strSex

30610     txtSurName = strSurName
30620     txtForeName = strForeName

30630     txtSex = strSex
          'Call DemographicsUniLabNoSelect(Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, txtSex, txtChart, txtLabNo)

       
30640     Exit Sub

       
txtForeName_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

30650     intEL = Erl
30660     strES = Err.Description
30670     LogError "frmEditSemen", "txtForeName_LostFocus", intEL, strES

End Sub

Private Sub txtsampleid_LostFocus()

30680     txtSampleID = Format$(Val(txtSampleID))
30690     If txtSampleID = 0 Then Exit Sub

30700     GetSampleIDWithOffset
          '40    txtLabNo = Val(FndMaxID("demographics", "LabNo", "")) + 1

30710     LoadAllDetails

30720     cmdSaveHold.Enabled = False
30730     cmdSave.Enabled = False

End Sub

Private Sub tSampleTime_KeyPress(KeyAscii As Integer)

30740     cmdSaveHold.Enabled = True
30750     cmdSave.Enabled = True

End Sub


Private Sub txtSex_Change()

30760     lblSex = txtSex

End Sub

Private Sub txtsex_Click()

30770     Select Case Trim$(txtSex)
              Case "": txtSex = "Male"
30780         Case "Male": txtSex = "Female"
30790         Case "Female": txtSex = ""
30800         Case Else: txtSex = ""
30810     End Select

30820     cmdSaveHold.Enabled = True
30830     cmdSave.Enabled = True

End Sub


Private Sub txtsex_KeyPress(KeyAscii As Integer)

30840     KeyAscii = 0
30850     txtsex_Click

End Sub


Private Sub txtSemenComment_KeyPress(KeyAscii As Integer)

30860     cmdSave.Enabled = True
30870     cmdSaveHold.Enabled = True

End Sub

Private Sub txtDemographicComment_KeyPress(KeyAscii As Integer)

30880     cmdSaveHold.Enabled = True
30890     cmdSave.Enabled = True

End Sub

Private Sub udMotility_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

          '10    If Trim(txtMotility(0)) = "" And _
          '         Trim(txtMotility(1)) <> "" And _
          '         Trim(txtMotility(2)) <> "" Then
          '20      txtMotility(0) = Format(100 - (Val(txtMotility(1)) + Val(txtMotility(2))))
          '30    End If
          '
          '40    If Trim(txtMotility(0)) <> "" And _
          '         Trim(txtMotility(1)) = "" And _
          '         Trim(txtMotility(2)) <> "" Then
          '50      txtMotility(1) = Format(100 - (Val(txtMotility(0)) + Val(txtMotility(2))))
          '60    End If
          '
          '70    If Trim(txtMotility(0)) <> "" And _
          '         Trim(txtMotility(1)) <> "" And _
          '         Trim(txtMotility(2)) = "" Then
          '80      txtMotility(2) = Format(100 - (Val(txtMotility(0)) + Val(txtMotility(1))))
          '90    End If

30900     cmdSave.Enabled = True
30910     cmdSaveHold.Enabled = True

End Sub



Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

30920     pBar = 0

30930     GetSampleIDWithOffset

30940     LoadAllDetails

30950     cmdSaveHold.Enabled = False
30960     cmdSave.Enabled = False

End Sub



Public Property Let PrintToPrinter(ByVal strNewValue As String)

30970     pPrintToPrinter = strNewValue

End Property
Public Property Get PrintToPrinter() As String

30980     PrintToPrinter = pPrintToPrinter

End Property

Private Sub LockSemenControls(Enable As Boolean)

30990     On Error GoTo LockSemenControls_Error

31000     Frame2.Enabled = Enable
31010     Frame10(0).Enabled = Enable
31020     Frame12.Enabled = Enable
31030     Frame11.Enabled = Enable
31040     Frame9.Enabled = Enable
31050     Frame14.Enabled = Enable

31060     Exit Sub

LockSemenControls_Error:

          Dim strES As String
          Dim intEL As Integer

31070     intEL = Erl
31080     strES = Err.Description
31090     LogError "frmEditSemen", "LockSemenControls", intEL, strES

End Sub

Private Sub LabNoUpdatePrvColor()
31100     On Error GoTo LabNoUpdatePrv_Error


          'If UCase(LabNoUpdatePrviousData) = UCase(txtSurName & txtForeName & txtDoB) Then
31110     If LabNoUpdatePrviousData = "1" Then
31120         txtLabNo.BackColor = vbGreen
              '40        lAddWardGP = FindLatestAddress(txtChart, Trim$(UCase$(txtSurName & " " & txtForeName)), txtDoB, Left$(txtSex, 1), txtLabNo)
31130     Else
              'LabNoUpdatePrviousData = ""
31140         txtLabNo.BackColor = vbRed
31150     End If


31160     Exit Sub


LabNoUpdatePrv_Error:

          Dim strES As String
          Dim intEL As Integer

31170     intEL = Erl
31180     strES = Err.Description
31190     LogError "frmEditAll", "LabNoUpdatePrv", intEL, strES
End Sub

Private Sub LabNoUpdatePrvData(ChartNo As String, Name As String, DoB As String, Sex As String, LabNo As String)
31200     On Error GoTo LabNoUpdatePrvData_Error
          Dim sql As String
31210     If txtLabNo.BackColor = vbGreen Then
31220         sql = "UPDATE demographics "
31230         sql = sql & " SET PatName ='" & Name & "'"
31240         sql = sql & ",  DoB ='" & DoB & "'"
31250         sql = sql & ",  Sex ='" & Sex & "'"
31260         sql = sql & ",  LabNo ='" & LabNo & "'"
31270         sql = sql & ",  Chart ='" & ChartNo & "'"
31280         sql = sql & " WHERE "
31290         sql = sql & " UPPER(PatName) ='" & UCase(Name) & "'"
31300         sql = sql & " AND DoB ='" & DoB & "'"
31310         sql = sql & " AND UPPER(Sex) ='" & UCase(Sex) & "'"
31320         sql = sql & " AND UPPER(Chart) ='" & UCase(ChartNo) & "'"

31330         Cnxn(0).Execute sql
31340     End If

31350     Exit Sub


LabNoUpdatePrvData_Error:

          Dim strES As String
          Dim intEL As Integer

31360     intEL = Erl
31370     strES = Err.Description
31380     LogError "frmEditAll", "LabNoUpdatePrvData", intEL, strES, sql

End Sub

Private Function FindLatestAddress(ChartNo As String, Name As String, DoB As String, Sex As String, LabNo As String) As String
          Dim sql As String

31390     On Error GoTo FindLatestAddress_Error

          Dim tb As New ADODB.Recordset
31400     sql = "Select Addr0 from demographics  "
31410     sql = sql & " WHERE "
31420     sql = sql & " UPPER(PatName) ='" & AddTicks(UCase(Name)) & "'"
31430     sql = sql & " AND DoB ='" & DoB & "'"
31440     sql = sql & " AND UPPER(Sex) ='" & UCase(Sex) & "'"
31450     sql = sql & " AND UPPER(Chart) ='" & UCase(ChartNo) & "'"
31460     sql = sql & " ORDER BY DateTimeDemographics DESC  "
31470     Set tb = New Recordset
31480     RecOpenServer 0, tb, sql

31490     If Not tb.EOF Then
31500         FindLatestAddress = tb!Addr0
31510     End If


31520     Exit Function


FindLatestAddress_Error:

          Dim strES As String
          Dim intEL As Integer

31530     intEL = Erl
31540     strES = Err.Description
31550     LogError "frmEditAll", "FindLatestAddress", intEL, strES, sql
End Function

Private Function DemographicsUniLabNoSelect(PatName As String, DoB As String, Sex As String, Chart As String, LabNo As String) As Double

31560     On Error GoTo DemographicsUniLabNoSelect_Error
          Dim sql As String
          Dim tb As New ADODB.Recordset
31570     If PatName = "" Or DoB = "" Or Sex = "" Then
31580         Exit Function
31590     End If


31600     sql = "select Top 1 ISNULL(LabNo,0) as LabNo  from DemographicsUniLabNo As D  " & _
              " WHERE ISNULL(LabNo,0)  <> 0 AND  D.PatName='" & AddTicks(PatName) & "' AND DoB = '" & Format(DoB, "dd/MMM/yyyy") & "'" & _
              " ORDER BY DateTimeOfRecord DESC "
              
31610     Set tb = New Recordset
31620     RecOpenClient 0, tb, sql

31630     If tb.EOF = False Then
31640         DemographicsUniLabNoSelect = tb!LabNo
31650     Else
31660         LabNo = Val(FndMaxID("demographics", "LabNo", ""))
31670         Call DemographicsUniLabNoInsertValues("", UserName, PatName, DoB, Sex, Chart, LabNo)
31680         DemographicsUniLabNoSelect = LabNo
31690     End If

31700     txtLabNo = DemographicsUniLabNoSelect

31710     Exit Function


DemographicsUniLabNoSelect_Error:

          Dim strES As String
          Dim intEL As Integer

31720     intEL = Erl
31730     strES = Err.Description
31740     LogError "frmEditSemen", "DemographicsUniLabNoSelect", intEL, strES
End Function


