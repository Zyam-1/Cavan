VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmxmatch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transfusion Laboratory"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9480
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer timFetchSampleID 
      Interval        =   2000
      Left            =   12270
      Top             =   -120
   End
   Begin VB.CommandButton cmdIDpanels 
      BackColor       =   &H008080FF&
      Caption         =   "ID Panels"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   4185
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdViewReports 
      Caption         =   "View Reports"
      Height          =   855
      Left            =   8460
      Picture         =   "fcdrAllocMonaghan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   5400
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Electronic X-Matching"
      Height          =   2460
      Left            =   6510
      TabIndex        =   103
      Top             =   1545
      Width           =   4905
      Begin VB.CommandButton cmdViewEligibility 
         Caption         =   "View"
         Height          =   675
         Left            =   3540
         Picture         =   "fcdrAllocMonaghan.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label lblEligibleTime 
         Alignment       =   2  'Center
         Caption         =   "Eligible for next 12 hours 10 minutes"
         Height          =   195
         Left            =   150
         TabIndex        =   122
         Top             =   2100
         Visible         =   0   'False
         Width           =   4425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vision Data"
         Height          =   195
         Index           =   2
         Left            =   2790
         TabIndex        =   119
         Top             =   1110
         Width           =   810
      End
      Begin VB.Image imgEligible 
         Height          =   225
         Index           =   6
         Left            =   2550
         Picture         =   "fcdrAllocMonaghan.frx":194C
         Top             =   1095
         Width           =   210
      End
      Begin VB.Label lblForced 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Manually Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   150
         TabIndex        =   115
         Top             =   1440
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Previously Eligible"
         Height          =   195
         Index           =   1
         Left            =   2790
         TabIndex        =   113
         Top             =   870
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No Adverse Reactions"
         Height          =   195
         Index           =   1
         Left            =   2790
         TabIndex        =   112
         Top             =   630
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Previous A/B Screens"
         Height          =   195
         Index           =   1
         Left            =   2790
         TabIndex        =   111
         Top             =   405
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Current A/B Screen"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   110
         Top             =   870
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Previous Group agreement"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   109
         Top             =   630
         Width           =   1890
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Previous Samples"
         Height          =   195
         Left            =   1050
         TabIndex        =   108
         Top             =   390
         Width           =   1260
      End
      Begin VB.Image imgEligible 
         Height          =   225
         Index           =   5
         Left            =   2550
         Picture         =   "fcdrAllocMonaghan.frx":1C22
         Top             =   870
         Width           =   210
      End
      Begin VB.Image imgEligible 
         Height          =   225
         Index           =   4
         Left            =   2550
         Picture         =   "fcdrAllocMonaghan.frx":1EF8
         Top             =   630
         Width           =   210
      End
      Begin VB.Image imgEligible 
         Height          =   225
         Index           =   3
         Left            =   2550
         Picture         =   "fcdrAllocMonaghan.frx":21CE
         Top             =   390
         Width           =   210
      End
      Begin VB.Image imgEligible 
         Height          =   225
         Index           =   2
         Left            =   2340
         Picture         =   "fcdrAllocMonaghan.frx":24A4
         Top             =   870
         Width           =   210
      End
      Begin VB.Image imgEligible 
         Height          =   225
         Index           =   1
         Left            =   2340
         Picture         =   "fcdrAllocMonaghan.frx":277A
         Top             =   630
         Width           =   210
      End
      Begin VB.Image imgEligible 
         Height          =   225
         Index           =   0
         Left            =   2340
         Picture         =   "fcdrAllocMonaghan.frx":2A50
         Top             =   390
         Width           =   210
      End
      Begin VB.Label lblEligible 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Eligible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1800
         TabIndex        =   105
         Top             =   1440
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblNotEligible 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Not Eligible"
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
         Height          =   315
         Left            =   1800
         TabIndex        =   104
         Top             =   1440
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdOrderAutovue 
      Caption         =   "Order"
      Height          =   855
      Left            =   12660
      Picture         =   "fcdrAllocMonaghan.frx":2D26
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   5400
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "Confirm Details"
      Height          =   855
      Left            =   7530
      Picture         =   "fcdrAllocMonaghan.frx":3DA8
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   5400
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   1005
      Left            =   11490
      Picture         =   "fcdrAllocMonaghan.frx":40B2
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   1200
      Width           =   885
   End
   Begin VB.CommandButton cmdLab 
      Caption         =   "Lab"
      Height          =   315
      Left            =   12570
      TabIndex        =   85
      Top             =   1530
      Width           =   915
   End
   Begin VB.CommandButton bgenotype 
      Appearance      =   0  'Flat
      Caption         =   "&Genotype"
      Height          =   315
      Left            =   12570
      TabIndex        =   84
      Top             =   1890
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   9630
      Picture         =   "fcdrAllocMonaghan.frx":44F4
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "bCancel"
      Top             =   5400
      Width           =   885
   End
   Begin VB.CommandButton bHold 
      Caption         =   "Save && &Hold"
      Enabled         =   0   'False
      Height          =   855
      Left            =   11490
      Picture         =   "fcdrAllocMonaghan.frx":5576
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   5400
      Width           =   885
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save F2"
      Enabled         =   0   'False
      Height          =   855
      Left            =   10545
      Picture         =   "fcdrAllocMonaghan.frx":65F8
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5385
      Width           =   885
   End
   Begin VB.CommandButton btnprint 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Print Label"
      Height          =   1005
      Left            =   11490
      Picture         =   "fcdrAllocMonaghan.frx":767A
      Style           =   1  'Graphical
      TabIndex        =   81
      Tag             =   "bprintlabels"
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton bPrintForm 
      Caption         =   "Print &Form"
      Height          =   975
      Left            =   12570
      Picture         =   "fcdrAllocMonaghan.frx":7ABC
      Style           =   1  'Graphical
      TabIndex        =   80
      Tag             =   "bprintform"
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdExternal 
      Caption         =   "Enter External Notes"
      Height          =   855
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   5400
      Width           =   885
   End
   Begin VB.Frame Frame8 
      Caption         =   "Cross Match"
      Height          =   2715
      Index           =   0
      Left            =   60
      TabIndex        =   65
      Top             =   6735
      Width           =   13395
      Begin VB.CommandButton bPrepare 
         Caption         =   "Electronic Issue"
         Height          =   315
         Index           =   1
         Left            =   8055
         TabIndex        =   120
         Top             =   150
         Width           =   1455
      End
      Begin VB.CommandButton cmdIssueToUnknown 
         Caption         =   "Issue to Unknown"
         Height          =   315
         Left            =   9570
         TabIndex        =   102
         Top             =   150
         Width           =   1635
      End
      Begin VB.CommandButton cmdSuggestUnits 
         Caption         =   "Suggest Units"
         Height          =   315
         Left            =   11400
         TabIndex        =   101
         Top             =   150
         Width           =   1275
      End
      Begin VB.CommandButton cmdShowHistory 
         Caption         =   "Show Current"
         Height          =   315
         Left            =   1110
         TabIndex        =   69
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton bIssueBatch 
         Caption         =   "Issue Batched Product"
         Height          =   315
         Left            =   3870
         TabIndex        =   67
         Top             =   150
         Width           =   2025
      End
      Begin VB.CommandButton bPrepare 
         Caption         =   "Manual Prepare"
         Height          =   315
         Index           =   0
         Left            =   6060
         TabIndex        =   66
         Top             =   150
         Width           =   1845
      End
      Begin MSFlexGridLib.MSFlexGrid gXmatch 
         Height          =   2145
         Left            =   90
         TabIndex        =   68
         Top             =   480
         Width           =   13245
         _ExtentX        =   23363
         _ExtentY        =   3784
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         FormatString    =   $"fcdrAllocMonaghan.frx":7EFE
      End
   End
   Begin VB.CommandButton bPrintDAT 
      Caption         =   "Print DAT"
      Height          =   495
      Left            =   12840
      TabIndex        =   63
      Top             =   2688
      Width           =   585
   End
   Begin VB.Frame Frame6 
      Height          =   6705
      Left            =   60
      TabIndex        =   30
      Top             =   30
      Width           =   6405
      Begin VB.CommandButton cmdUpDown 
         Caption         =   "+"
         Height          =   285
         Index           =   0
         Left            =   1125
         TabIndex        =   124
         Top             =   495
         Width           =   285
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
         Height          =   285
         Index           =   1
         Left            =   750
         TabIndex        =   123
         Top             =   495
         Width           =   285
      End
      Begin VB.TextBox txtSurname 
         Height          =   240
         Left            =   5100
         MaxLength       =   30
         TabIndex        =   118
         Tag             =   "Name"
         Top             =   600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtForname 
         Height          =   240
         Left            =   5100
         MaxLength       =   30
         TabIndex        =   117
         Tag             =   "Name"
         Top             =   810
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame FramePP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4485
         TabIndex        =   96
         Top             =   3720
         Visible         =   0   'False
         Width           =   1875
         Begin VB.OptionButton optPubPri 
            Caption         =   "Private"
            Height          =   195
            Index           =   1
            Left            =   945
            TabIndex        =   98
            Top             =   135
            Width           =   810
         End
         Begin VB.OptionButton optPubPri 
            Alignment       =   1  'Right Justify
            Caption         =   "Public"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   97
            Top             =   135
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdLockDemo 
         Caption         =   "Lock"
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
         Left            =   3930
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   135
         Width           =   1035
      End
      Begin VB.TextBox txtSampleTime 
         Height          =   285
         Left            =   5310
         TabIndex        =   14
         Top             =   5280
         Width           =   1005
      End
      Begin VB.TextBox tComment 
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
         Left            =   750
         TabIndex        =   91
         Top             =   3180
         Width           =   3615
      End
      Begin VB.CommandButton cmdChangeHistory 
         BackColor       =   &H0000FFFF&
         Caption         =   "C/H"
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
         Left            =   -240
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "View Change History"
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox tTypenex 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "### ####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   6153
            SubFormatType   =   0
         EndProperty
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
         Left            =   3900
         TabIndex        =   74
         Top             =   795
         Width           =   1095
      End
      Begin VB.TextBox txtNOPAS 
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
         Left            =   750
         TabIndex        =   73
         Top             =   225
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox tAandE 
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
         Left            =   3900
         TabIndex        =   71
         Top             =   495
         Width           =   1095
      End
      Begin VB.ComboBox cGP 
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
         Left            =   750
         TabIndex        =   11
         Text            =   "cGP"
         Top             =   4140
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
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
         Index           =   3
         Left            =   750
         MaxLength       =   50
         TabIndex        =   37
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
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
         Index           =   2
         Left            =   750
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2580
         Width           =   3615
      End
      Begin VB.TextBox tMaiden 
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
         Left            =   750
         MaxLength       =   20
         TabIndex        =   36
         Top             =   1665
         Width           =   3615
      End
      Begin VB.ComboBox cSpecial 
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
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   750
         TabIndex        =   20
         Text            =   "cSpecial"
         Top             =   5040
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
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
         Left            =   750
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
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
         Left            =   750
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1965
         Width           =   3615
      End
      Begin VB.ComboBox cWard 
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
         IntegralHeight  =   0   'False
         Left            =   750
         TabIndex        =   9
         Text            =   "cWard"
         Top             =   3480
         Width           =   3615
      End
      Begin VB.ComboBox cProcedure 
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
         IntegralHeight  =   0   'False
         Left            =   750
         TabIndex        =   19
         Text            =   "cProcedure"
         Top             =   4740
         Width           =   3615
      End
      Begin VB.ComboBox cConditions 
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
         IntegralHeight  =   0   'False
         Left            =   750
         TabIndex        =   12
         Text            =   "cConditions"
         Top             =   4440
         Width           =   3615
      End
      Begin VB.ComboBox cClinician 
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
         IntegralHeight  =   0   'False
         Left            =   750
         TabIndex        =   10
         Text            =   "cClinician"
         Top             =   3810
         Width           =   3615
      End
      Begin VB.TextBox tLabNum 
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
         Left            =   750
         MaxLength       =   14
         TabIndex        =   0
         Top             =   795
         Width           =   2175
      End
      Begin VB.TextBox txtName 
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
         Left            =   750
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1365
         Width           =   3615
      End
      Begin VB.TextBox txtChart 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   1
         Top             =   495
         Width           =   1155
      End
      Begin VB.CommandButton bhistory 
         Caption         =   "History"
         Height          =   315
         Left            =   5070
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1050
         Width           =   1245
      End
      Begin VB.TextBox tedd 
         Height          =   285
         Left            =   5310
         MaxLength       =   8
         TabIndex        =   34
         Top             =   4650
         Width           =   1005
      End
      Begin VB.TextBox tAge 
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
         Left            =   4980
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1680
         Width           =   1125
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5070
         TabIndex        =   33
         Tag             =   "bsearch"
         Top             =   330
         Width           =   1245
      End
      Begin VB.TextBox tDoB 
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
         Left            =   4980
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1380
         Width           =   1125
      End
      Begin VB.TextBox txtSampleDate 
         Height          =   285
         Left            =   5310
         MaxLength       =   10
         TabIndex        =   13
         Top             =   4980
         Width           =   1005
      End
      Begin VB.TextBox tSampleComment 
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
         Left            =   750
         MaxLength       =   50
         TabIndex        =   32
         Top             =   5340
         Width           =   3615
      End
      Begin ComCtl2.UpDown udLabNum 
         Height          =   225
         Left            =   -600
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         _Version        =   327681
         Value           =   1
         OrigLeft        =   810
         OrigTop         =   120
         OrigRight       =   1800
         OrigBottom      =   345
         Increment       =   0
         Max             =   99999
         Min             =   1
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtTimeRxd 
         Height          =   285
         Left            =   4890
         TabIndex        =   16
         Top             =   6090
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   183173123
         UpDown          =   -1  'True
         CurrentDate     =   37889.9993055556
      End
      Begin MSComCtl2.DTPicker dtDateRxd 
         Height          =   285
         Left            =   4680
         TabIndex        =   15
         Top             =   5790
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   183173123
         CurrentDate     =   37889
      End
      Begin VB.Label lblPrevAdverse 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Patient with previous Adverse Reaction!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Left            =   60
         TabIndex        =   106
         Top             =   5700
         Visible         =   0   'False
         Width           =   4545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Sample Time"
         Height          =   195
         Index           =   0
         Left            =   4380
         TabIndex        =   93
         Top             =   5310
         Width           =   915
      End
      Begin VB.Image imgSquareTick 
         Height          =   225
         Left            =   5310
         Picture         =   "fcdrAllocMonaghan.frx":7FFA
         Top             =   6420
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgSquare 
         Height          =   225
         Left            =   5550
         Picture         =   "fcdrAllocMonaghan.frx":82D0
         Top             =   6420
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgSquareCross 
         Height          =   225
         Left            =   5760
         Picture         =   "fcdrAllocMonaghan.frx":85A6
         Top             =   6420
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Date/Time Received"
         Height          =   195
         Left            =   4680
         TabIndex        =   83
         Top             =   5580
         Width           =   1500
      End
      Begin VB.Image imgUseTime 
         Height          =   225
         Left            =   5790
         Picture         =   "fcdrAllocMonaghan.frx":887C
         Top             =   6120
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label13 
         Caption         =   "TYPENEX"
         Height          =   195
         Left            =   3120
         TabIndex        =   76
         Top             =   825
         Width           =   765
      End
      Begin VB.Label lblNOPAS 
         AutoSize        =   -1  'True
         Caption         =   "NOPAS"
         Height          =   195
         Left            =   150
         TabIndex        =   75
         Top             =   270
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblAandE 
         AutoSize        =   -1  'True
         Caption         =   "A/E"
         Height          =   195
         Left            =   3600
         TabIndex        =   72
         Top             =   525
         Width           =   285
      End
      Begin VB.Label lSex 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Left            =   4980
         TabIndex        =   5
         Top             =   1980
         Width           =   1125
      End
      Begin VB.Label lblChartNumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monaghan Chart #"
         Height          =   285
         Left            =   2130
         TabIndex        =   64
         ToolTipText     =   "Click to change Location"
         Top             =   225
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GP"
         Height          =   195
         Left            =   420
         TabIndex        =   62
         Top             =   4200
         Width           =   225
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   570
         TabIndex        =   61
         Top             =   2940
         Width           =   90
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   570
         TabIndex        =   60
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   570
         TabIndex        =   59
         Top             =   2340
         Width           =   90
      End
      Begin VB.Label lblMaiden 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "M.Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   58
         Top             =   1725
         Width           =   600
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Remark"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   90
         TabIndex        =   57
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Spec"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   27
         Left            =   270
         TabIndex        =   56
         Top             =   5100
         Width           =   375
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Addr 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   150
         TabIndex        =   55
         Top             =   2025
         Width           =   465
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Proc"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   29
         Left            =   330
         TabIndex        =   54
         Top             =   4800
         Width           =   330
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cond"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   30
         Left            =   270
         TabIndex        =   53
         Top             =   4500
         Width           =   375
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Clin"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   31
         Left            =   390
         TabIndex        =   52
         Top             =   3870
         Width           =   255
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sample #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   32
         Left            =   30
         TabIndex        =   51
         Top             =   855
         Width           =   675
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   33
         Left            =   270
         TabIndex        =   50
         Top             =   3540
         Width           =   390
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   34
         Left            =   210
         TabIndex        =   49
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   4680
         TabIndex        =   48
         Top             =   2010
         Width           =   270
      End
      Begin VB.Label lblChartTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MRN"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1860
         TabIndex        =   47
         Top             =   525
         Width           =   375
      End
      Begin VB.Label lblEDD 
         AutoSize        =   -1  'True
         Caption         =   "EDD"
         Height          =   195
         Left            =   4950
         TabIndex        =   46
         Top             =   4710
         Width           =   345
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   4650
         TabIndex        =   45
         Top             =   1740
         Width           =   285
      End
      Begin VB.Label l 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   4470
         TabIndex        =   44
         Top             =   1440
         Width           =   450
      End
      Begin VB.Image iprevious 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4500
         Top             =   2310
         Width           =   1605
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Sample Date"
         Height          =   195
         Index           =   0
         Left            =   4350
         TabIndex        =   43
         Top             =   5025
         Width           =   945
      End
      Begin VB.Label lKnownAntibody 
         BackColor       =   &H000000FF&
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
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   750
         TabIndex        =   42
         Top             =   1080
         Visible         =   0   'False
         Width           =   4260
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Sample Comment"
         Height          =   360
         Index           =   0
         Left            =   30
         TabIndex        =   41
         Top             =   5310
         Width           =   660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lchk 
         AutoSize        =   -1  'True
         Caption         =   "Checked By:"
         Height          =   195
         Left            =   4470
         TabIndex        =   40
         Top             =   4125
         Width           =   915
      End
      Begin VB.Label lblgrpchecker 
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
         Height          =   315
         Left            =   4500
         TabIndex        =   39
         Top             =   4335
         Width           =   1815
      End
      Begin VB.Label lInfo 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This sample was processed more than seventy-two hours ago. Transfusion time should be checked."
         ForeColor       =   &H000000FF&
         Height          =   1200
         Left            =   4500
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   1755
      End
   End
   Begin VB.TextBox tident 
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
      Left            =   6510
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   4530
      Width           =   4905
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   225
      Left            =   6450
      TabIndex        =   28
      Top             =   6360
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame FrLockable 
      Caption         =   "Grouping"
      Height          =   1425
      Index           =   0
      Left            =   6510
      TabIndex        =   23
      Top             =   30
      Width           =   4875
      Begin VB.ComboBox cmbKell 
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
         Left            =   855
         TabIndex        =   89
         Top             =   930
         Width           =   1035
      End
      Begin VB.ComboBox lstRG 
         Appearance      =   0  'Flat
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
         Left            =   2265
         TabIndex        =   78
         Text            =   "lstRG"
         Top             =   510
         Width           =   765
      End
      Begin VB.ComboBox lstfg 
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
         Left            =   855
         TabIndex        =   24
         Top             =   510
         Width           =   1035
      End
      Begin VB.Label lblGroupDisagree 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Historical Grouping Disagreement"
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
         Height          =   450
         Left            =   2280
         TabIndex        =   114
         Top             =   945
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Kell"
         Height          =   195
         Left            =   555
         TabIndex        =   88
         Top             =   990
         Width           =   255
      End
      Begin VB.Label lblSuggestRG 
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
         Left            =   2265
         TabIndex        =   79
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lDontMatch 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FG / RG Don't Match"
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
         Height          =   615
         Left            =   3075
         TabIndex        =   77
         Top             =   210
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblsuggestfg 
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
         Left            =   855
         TabIndex        =   27
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Suggest"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   26
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Report"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   25
         Top             =   570
         Width           =   480
      End
   End
   Begin VB.Frame FrLockable 
      Height          =   2565
      Index           =   2
      Left            =   11490
      TabIndex        =   22
      Top             =   2700
      Width           =   1965
      Begin MSFlexGridLib.MSFlexGrid gDAT 
         CausesValidation=   0   'False
         Height          =   1815
         Left            =   90
         TabIndex        =   87
         Top             =   660
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   3
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   0
         BorderStyle     =   0
         FormatString    =   "<DAT              |^Pos|^Neg"
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
      Begin VB.Label lblOrderDAT 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Order DAT"
         Height          =   495
         Left            =   780
         TabIndex        =   90
         Top             =   120
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.CommandButton bviewwl 
      Appearance      =   0  'Flat
      Caption         =   "&Work Lists"
      Height          =   345
      Left            =   12570
      TabIndex        =   21
      Tag             =   "viewwl"
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "AB Report"
      Height          =   195
      Left            =   6540
      TabIndex        =   100
      Top             =   4290
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   14880
      Picture         =   "fcdrAllocMonaghan.frx":8B52
      Top             =   3420
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmxmatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim labnumberfound As String

Dim SearchCriteria As String

Dim Loading As Boolean

Private mNewRecord As Boolean

Private blnShowingBox As Boolean

Private UpDownDirection As Integer

Public g_SampleID As String


Private Function FlagMessage(ByVal strType As String, _
                             ByVal Historical As String, _
                             ByVal Current As String) _
                             As Boolean
'Returns True to reject

    Dim s As String

10  If Trim$(Historical) = "" Then Historical = "<Blank>"
20  If Trim$(Current) = "" Then Current = "<Blank>"
    If Trim$(Historical) = "" Then
        FlagMessage = True
        Exit Function
    End If
30  s = "Patients " & strType & " has changed!" & vbCrLf & _
        "Was '" & Historical & "'" & vbCrLf & _
        "Now '" & Current & "'" & vbCrLf & _
        "To accept this change, Press 'OK'"

40  FlagMessage = iMsg(s, vbCritical + vbOKCancel, "Critical Warning") = vbCancel

End Function

Private Sub GetPrevious()

    Dim sql As String
    Dim tb As Recordset
    Dim TotChart As Long
    Dim TotName As Long
    Dim Tot As Long

10  On Error GoTo GetPrevious_Error

20  If Trim$(txtChart) <> "" Then
30      sql = "SELECT count (*) as tot FROM patientdetails WHERE " & _
              "patnum = '" & AddTicks(txtChart) & "'"
40      Set tb = New Recordset
50      RecOpenServerBB 0, tb, sql
60      TotChart = tb!Tot
70  End If

80  If Trim$(txtName) <> "" Then
90      sql = "SELECT count (*) as tot FROM patientdetails WHERE " & _
              "Name = '" & AddTicks(txtName) & "'"
100     Set tb = New Recordset
110     RecOpenServerBB 0, tb, sql
120     TotName = tb!Tot
130 End If

140 If TotChart > TotName Then
150     Tot = TotChart
160 Else
170     Tot = TotName
180 End If
190 If Tot <> 0 Then
200     cmdSearch.Caption = "Search (" & Format$(Tot) & ")"
210 Else
220     cmdSearch.Caption = "Search"
230 End If

240 Exit Sub

GetPrevious_Error:

    Dim strES As String
    Dim intEL As Integer

250 intEL = Erl
260 strES = Err.Description
270 LogError "frmxmatch", "GetPrevious", intEL, strES, sql


End Sub

Private Function IsAntibodyHistoryOK() As Boolean

    Dim sqlBase As String
    Dim sql As String
    Dim tb As Recordset
    Dim s As String

10  On Error GoTo IsAntibodyHistoryOK_Error

20  IsAntibodyHistoryOK = True

30  sqlBase = "SELECT AIDR, PatNum, Name, DoB, LabNumber, DateTime FROM PatientDetails WHERE " & _
              "AIDR LIKE '%pos%' "

40  If Trim$(txtChart) <> "" Then
50      sql = sqlBase & "AND PatNum = '" & Trim$(txtChart) & "'"
60      Set tb = New Recordset
70      RecOpenServerBB 0, tb, sql
80      If Not tb.EOF Then
90          s = "Patient with Chart Number " & txtChart & vbCrLf & _
                "Name " & tb!Name & vbCrLf & _
                "DoB " & tb!DoB & vbCrLf & _
                tb!AIDR & vbCrLf & _
                "(Lab SampleID " & tb!LabNumber & " )" & vbCrLf & _
                "Continue?"
100         Answer = iMsg(s, vbQuestion + vbYesNo, "Antibody Check", vbRed)
110         If TimedOut Then Unload Me: Exit Function
120         If Answer = vbNo Then
130             IsAntibodyHistoryOK = False
140         Else
150             LogReasonWhy "Positive Antibody History warning issued. (Chart)", "X"
160         End If
170     End If
180     Exit Function
190 End If
200 If Trim$(txtName) <> "" Then
210     If Left$(UCase$(txtName), 2) = "O'" Then
220         sql = sqlBase & "AND Name LIKE '" & AddTicks(txtName) & "%' "
230     Else
240         sql = sqlBase & "AND Name LIKE '" & txtName & "%' "
    'sql = sqlBase & "AND soundex(Name) = soundex('" & AddTicks(txtName) & "') "
250     End If
260     Set tb = New Recordset
270     RecOpenServerBB 0, tb, sql
280     If Not tb.EOF Then
290         s = "Patient Named " & tb!Name & vbCrLf & _
                "Chart Number " & tb!Patnum & vbCrLf & _
                "DoB " & tb!DoB & vbCrLf & _
                tb!AIDR & vbCrLf & _
                "(Lab SampleID " & tb!LabNumber & " )" & vbCrLf & _
                "Continue?"
300         Answer = iMsg(s, vbQuestion + vbYesNo, "Antibody Check", vbRed)
310         If TimedOut Then Unload Me: Exit Function
320         If Answer = vbNo Then
330             IsAntibodyHistoryOK = False
340         Else
350             LogReasonWhy "Positive Antibody History warning issued.(Name)", "X"
360         End If
370     End If
380 End If
390 If IsDate(tDoB) Then
400     sql = sqlBase & "AND DoB = '" & Format$(tDoB, "dd/MMM/yyyy") & "' "
410     Set tb = New Recordset
420     RecOpenServerBB 0, tb, sql
430     If Not tb.EOF Then
440         s = "Patient with Date of Birth " & tDoB & vbCrLf & _
                "Chart Number " & tb!Patnum & vbCrLf & _
                "Named " & tb!Name & vbCrLf & _
                tb!AIDR & vbCrLf & _
                "(Lab SampleID " & tb!LabNumber & " )" & vbCrLf & _
                "Continue?"
450         Answer = iMsg(s, vbQuestion + vbYesNo, "Antibody Check", vbRed)
460         If TimedOut Then Unload Me: Exit Function
470         If Answer = vbNo Then
480             IsAntibodyHistoryOK = False
490         Else
500             LogReasonWhy "Positive Antibody History warning issued.(DoB)", "X"
510         End If
520     End If
530 End If

540 Exit Function

IsAntibodyHistoryOK_Error:

    Dim strES As String
    Dim intEL As Integer

550 intEL = Erl
560 strES = Err.Description
570 LogError "frmxmatch", "IsAntibodyHistoryOK", intEL, strES, sql


End Function

Private Function IsChangeHistory() As Boolean

    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo IsChangeHistory_Error

20  sql = "select * from PatientDetailsAudit where " & _
          "labnumber = '" & tLabNum & "'"
30  Set tb = New Recordset
40  RecOpenClientBB 0, tb, sql
50  If tb.EOF Then
60      IsChangeHistory = False
70  Else
80      IsChangeHistory = True
90  End If

100 Exit Function

IsChangeHistory_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmxmatch", "IsChangeHistory", intEL, strES, sql


End Function

Public Sub LoadLabNumber()

    Dim sn As Recordset
    Dim sql As String
    Dim s As String
    Dim n As Integer
    Dim PrevGroup As String
    Dim PrevRH As String
    Dim FoundDAT As Boolean

10  On Error GoTo LoadLabNumber_Error

20  lInfo.Visible = False
30  lKnownAntibody.Visible = False  '(Nk)
40  cmdChangeHistory.Visible = False
50  If IsChangeHistory() Then
60      cmdChangeHistory.Visible = True
70  End If
80  lblPrevAdverse.Visible = False
90  cmdShowHistory.Caption = "Hide Current"
100 gXmatch.Visible = True

110 s = tLabNum

120 ClearDetails

130 cmdSave.Enabled = False: bHold.Enabled = False
140 tLabNum = s

150 mNewRecord = True

160 If Trim$(tLabNum) = "" Then
170     If Screen.ActiveControl.Tag = "viewwl" Then
180         bviewwl = True
190     ElseIf Screen.ActiveControl.Tag = "cancel" Then
200         cmdCancel = True
210     End If
220     Exit Sub
230 End If

240 sql = "select * from patientdetails where " & _
          "labnumber = '" & tLabNum & "'"
250 Set sn = New Recordset
260 RecOpenServerBB 0, sn, sql

270 If Not sn.EOF Then
280     LockDemographics True
290     txtName = sn!Name & ""
300     txtSurname = sn!PatSurName & ""
310     txtForname = sn!PatForeName & ""
320     txtChart = sn!Patnum & ""

330     lblGroupDisagree.Visible = False
340     If Not GroupHistoryOK() Then
350         lblGroupDisagree.Visible = True
360     End If

370     cmbKell = sn!Kell & ""

380     If Trim$(txtChart) <> "" Then
390         If CheckBadReaction(txtChart) Then
400             lblPrevAdverse.Visible = True
410         End If
420         lKnownAntibody.Caption = CheckPreviousABScreen(txtChart)
430         If lKnownAntibody.Caption <> "" Then
440             lKnownAntibody.Visible = True
450         End If
460     End If
470     tTypenex = sn!Typenex & ""
480     tAandE = sn!AandE & ""
490     txtNOPAS = sn!NOPAS & ""
500     mNewRecord = False
510     lblgrpchecker = sn!Checker & ""
520     lblgrpchecker.Enabled = lblgrpchecker = ""

530     Select Case Left$(sn!Sex & "", 1)
            Case "M": lSex = "Male"
540         Case "F": lSex = "Female"
550         Case "U": lSex = "Unknown"
560         Case Else: lSex = ""
570     End Select

580     tAddr(0) = sn!Addr1 & ""
590     tAddr(1) = sn!Addr2 & ""
600     tAddr(2) = sn!Addr3 & ""
610     tAddr(3) = sn!addr4 & ""
620     tMaiden = sn!maiden & ""
630     If Not IsNull(sn!DoB) Then
640         tDoB = Format(sn!DoB, "dd/mm/yyyy")
650     Else
660         tDoB = ""
670     End If
680     If Trim$(sn!Age & "") <> "" Then
690         tAge = sn!Age
700     Else
710         tAge = CalcAge(tDoB)
720     End If
730     cWard = sn!Ward & ""
740     cGP = sn!GP & ""
750     cClinician = sn!Clinician & ""

760     If Not IsNull(sn!IsPublic) Then
770         optPubPri(0).Value = sn!IsPublic    'Public
780         optPubPri(1).Value = Not (sn!IsPublic)    'Private
790     End If

800     cConditions = sn!Conditions & ""
810     cProcedure = sn!Procedure & ""
820     cSpecial = sn!specialprod & ""
    'grh2image trim$(left$(sn!PrevGroup & "", 2)), sn!previousrh & ""
830     tComment = StripComment(sn!Comment & "")
840     lblsuggestfg = sn!fgsuggest & ""
850     tedd = sn!edd & ""
860     tSampleComment = sn!SampleComment & ""

870     lblSuggestRG = sn!rgSuggest & ""
880     lstRG = sn!rgroup & ""

890     lstfg = sn!fGroup & ""

900     tident = sn!AIDR & ""

910     If IsNull(sn!SampleDate) Then
920         txtSampleDate = ""
930         txtSampleTime = ""
940     ElseIf Not IsDate(sn!SampleDate & "") Then
950         txtSampleDate = Format(Now, "dd/mm/yyyy")
960         txtSampleTime = Format(Now, "HH:nn:ss")
970     Else
980         txtSampleDate = Format(sn!SampleDate, "dd/mm/yyyy")
990         If Format(sn!SampleDate, "HH:mm") <> "00:00" Then
1000            txtSampleTime = Format(sn!SampleDate, "HH:nn:ss")
1010        Else
1020            txtSampleTime = ""
1030        End If
1040    End If

1050    If IsNull(sn!DateReceived) Then
1060        dtDateRxd = Format(Now, "dd/MM/yyyy")
'1070        dtTimeRxd = "00:00"
1080        Set imgUseTime.Picture = imgSquareCross.Picture
1090        dtTimeRxd.Enabled = True
1100    Else
1110        dtDateRxd = Format(sn!DateReceived, "dd/MM/yyyy")
1120        dtTimeRxd = Format(TimeValue(sn!DateReceived), "HH:nn:ss")
1130        If Format(TimeValue(sn!DateReceived), "HH:mm") <> "00:00" Then
1140            Set imgUseTime.Picture = imgSquareTick.Picture
1150            dtTimeRxd.Enabled = True
1160        Else
1170            Set imgUseTime.Picture = imgSquareCross.Picture
1180            dtTimeRxd.Enabled = True
1190        End If
1200    End If

1210    FoundDAT = False
1220    For n = 0 To 11
1230        If Not IsNull(sn("DAT" & Format(n))) Then
1240            If sn("DAT" & Format(n)) Then
1250                gDAT.col = (n Mod 2) + 1
1260                If gDAT.Rows > (n \ 2) + 1 Then
1270                    gDAT.row = (n \ 2) + 1
1280                    Set gDAT.CellPicture = imgSquareCross.Picture
1290                    FoundDAT = True
1300                End If
1310            End If
1320        End If
1330    Next

1340    PrevGroup = GroupKnown(txtChart, tLabNum, txtName)
1350    lblGroupDisagree.Visible = False
1360    If PrevGroup <> "" Then
1370        If UCase$(Trim$(PrevGroup)) <> UCase$(Trim$(lstfg)) Then
1380            lblGroupDisagree.Visible = True
1390        End If
1400        PrevGroup = sn!PrevGroup & sn!previousrh & ""    ' GroupKnown(txtChart, tLabNum, txtName)
1410        PrevRH = ""
1420        If InStr(UCase$(PrevGroup), "POS") Then PrevRH = "+"
1430        If InStr(UCase$(PrevGroup), "NEG") Then PrevRH = "-"
1440        If InStr(PrevGroup, "-") Then PrevRH = "-"
1450        If InStr(PrevGroup, "+") Then PrevRH = "+"
1460        PrevGroup = Trim$(Left$(PrevGroup, 2))
1470        grh2image PrevGroup, PrevRH
1480    End If

1490    FillXMHistory

1500 Else
1510    LockDemographics False
1520    'iprevious.Picture = frmMain.ImageList1.ListImages("grPrev").Picture

1530    lstfg = ""

1540 End If

1550 If Not CheckExternalNotes("MRN", txtChart) Then
1560    CheckExternalNotes "AandE", tAandE
1570 End If

1580 cmdSave.Enabled = False: bHold.Enabled = False

1590 labnumberfound = tLabNum

1600 GetPrevious

1610 CheckIfPhoned

1620 ShowEligibility

1630 cmdIDpanels.Visible = isIDPanelPresent4SID(tLabNum)

1640 Exit Sub

LoadLabNumber_Error:

    Dim strES As String
    Dim intEL As Integer

1650 intEL = Erl
1660 strES = Err.Description
1670 LogError "frmxmatch", "LoadLabNumber", intEL, strES, sql

End Sub


Private Sub ShowEligibility()

    Dim EITime As String
    Dim SampleDateTime As String
    Dim n As Integer
    Dim EI As New ElectronicIssue
    Dim All_EI_Rules_Pass As Boolean

10  On Error GoTo ShowEligibility_Error

20  For n = 0 To 6
30      imgEligible(n).Visible = False
40  Next

50  lblEligible.Visible = False
60  lblNotEligible.Visible = False
70  lblEligibleTime.Visible = False
80  lblForced.Visible = False
90  All_EI_Rules_Pass = False

100 If IsDate(txtSampleDate) And IsDate(txtSampleTime) And Trim$(txtChart) <> "" Then
110     All_EI_Rules_Pass = True

120     For n = 0 To 6
130         imgEligible(n).Visible = True
140         imgEligible(n).Picture = imgSquare.Picture
150     Next

160     SampleDateTime = txtSampleDate & " " & txtSampleTime
170     EI.Chart = txtChart
180     EI.SampleDate = SampleDateTime
190     EI.SampleID = tLabNum
200     EI.Load

210     Select Case EI.PreviousSample
            Case 0: imgEligible(0).Picture = imgSquareCross.Picture: All_EI_Rules_Pass = False
220         Case 1: imgEligible(0).Picture = imgSquareTick.Picture
230         Case 2: imgEligible(0).Picture = imgSquare.Picture: All_EI_Rules_Pass = False
240     End Select

250     Select Case EI.PreviousGroupAgreement
            Case 0: imgEligible(1).Picture = imgSquareCross.Picture: All_EI_Rules_Pass = False
260         Case 1: imgEligible(1).Picture = imgSquareTick.Picture
270         Case 2: imgEligible(1).Picture = imgSquare.Picture: All_EI_Rules_Pass = False
280     End Select

290     Select Case EI.CurrentNegativeAB
            Case 0: imgEligible(2).Picture = imgSquareCross.Picture: All_EI_Rules_Pass = False
300         Case 1: imgEligible(2).Picture = imgSquareTick.Picture
310         Case 2: imgEligible(2).Picture = imgSquare.Picture: All_EI_Rules_Pass = False
320     End Select

330     Select Case EI.PreviousNegativeAB
            Case 0: imgEligible(3).Picture = imgSquareCross.Picture: All_EI_Rules_Pass = False
340         Case 1: imgEligible(3).Picture = imgSquareTick.Picture
350         Case 2: imgEligible(3).Picture = imgSquare.Picture: All_EI_Rules_Pass = False
360     End Select

370     Select Case EI.AdverseReactions
            Case 0: imgEligible(4).Picture = imgSquareTick.Picture
380         Case 1: imgEligible(4).Picture = imgSquareCross.Picture: All_EI_Rules_Pass = False
390         Case 2: imgEligible(4).Picture = imgSquare.Picture: All_EI_Rules_Pass = False
400     End Select

410     Select Case EI.PreviousSampleEligible
            Case 0: imgEligible(5).Picture = imgSquareCross.Picture: All_EI_Rules_Pass = False
420         Case 1: imgEligible(5).Picture = imgSquareTick.Picture
430         Case 2: imgEligible(5).Picture = imgSquare.Picture: All_EI_Rules_Pass = False
440     End Select

450     Select Case EI.ResultAbnormalFlags
            Case 0: imgEligible(6).Picture = imgSquareCross.Picture: All_EI_Rules_Pass = False
460         Case 1: imgEligible(6).Picture = imgSquareTick.Picture
470         Case 2: imgEligible(6).Picture = imgSquare.Picture: All_EI_Rules_Pass = False
480     End Select

490     lblEligible.Visible = False
500     lblNotEligible.Visible = True
510     lblEligibleTime.Visible = False

520     If (EI.ForcedEligible = 1 And Not All_EI_Rules_Pass) Or (EI.ForcedNotEligible = 1 And All_EI_Rules_Pass) Then
530         lblForced.Visible = True
540     End If
550 End If

560 If IsDate(txtSampleDate) And IsDate(txtSampleTime) Then
570     If EI.ForcedEligible = 1 Or (All_EI_Rules_Pass And Not (EI.ForcedNotEligible = 1)) Then
580         lblEligible.Visible = True
590     End If
600     lblEligibleTime.Visible = True
610     EITime = GetEITime(txtSampleDate, txtSampleTime, txtChart)
620     If IsDate(EITime) Then
630         lblEligibleTime.Caption = "Available until " & Format$(EITime, "dd/MM/yyyy HH:nn")
640     Else
650         lblEligibleTime.Caption = EITime
660     End If
670 End If

680 Exit Sub

ShowEligibility_Error:

    Dim strES As String
    Dim intEL As Integer

690 intEL = Erl
700 strES = Err.Description
710 LogError "frmxmatch", "ShowEligibility", intEL, strES

End Sub


Private Sub SaveAutomaticEligibility()

    Dim SampleDateTime As String
    Dim EI As New ElectronicIssue
    Dim sql As String

10  On Error GoTo SaveAutomaticEligibility_Error

20  If IsDate(txtSampleDate) And IsDate(txtSampleTime) And Trim$(txtChart) <> "" Then

30      SampleDateTime = txtSampleDate & " " & txtSampleTime
40      EI.Chart = txtChart
50      EI.SampleDate = SampleDateTime
60      EI.SampleID = tLabNum
70      EI.Load

    'IF Previous Samples = YES,  Previous Group agreement = YES, Current A/B Screen = YES, Previous A/B Screens = YES
    'No Adverse Reactions = YES, No Previous 'Not Eligible' = YES, Vision Data NOT modified = YES
    'A_E = "Automatically Eligible for Electronic Issue"
    'ANE = "Automatically NOT Eligible for Electronic Issue"

    'Previous code ----------------------------------------------------
    '80      If EI.PreviousSample = 1 And EI.PreviousGroupAgreement = 1 And EI.CurrentNegativeAB = 1 And EI.PreviousNegativeAB = 1 _
     '            And EI.AdverseReactions = 0 And EI.PreviousSampleEligible = 1 And EI.ResultAbnormalFlags = 1 Then 'All 7 Green ticks then mark sample as ELIGIBLE
    '           'Set as A_E (Automatically Eligible for Electronic Issue) where it was not Forcibly set previously
    '90          sql = "UPDATE PatientDetails SET Eligible4EI = 'A_E' WHERE (labnumber = '" & tLabNum & "') and (Eligible4EI is null or Eligible4EI = '' or Eligible4EI = 'A_E' or  Eligible4EI = 'ANE')"
    '100     ElseIf EI.PreviousGroupAgreement = 0 Or EI.CurrentNegativeAB = 0 Or EI.PreviousNegativeAB = 0 _
     '            Or EI.AdverseReactions = 1 Or EI.PreviousSampleEligible = 0 Or EI.ResultAbnormalFlags = 0 Then 'If any of rules 2 - 7 X(fail) then mark sample as  NOT ELIGIBLE
    '           'Set as ANE (Automatically NOT Eligible for Electronic Issue) where it was not Forcibly set previously
    '110         sql = "UPDATE PatientDetails SET Eligible4EI = 'ANE' WHERE (labnumber = '" & tLabNum & "') and (Eligible4EI is null or Eligible4EI = '' or Eligible4EI = 'A_E' or  Eligible4EI = 'ANE')"
    '        Else
    '            sql = "UPDATE PatientDetails SET Eligible4EI = '' WHERE (labnumber = '" & tLabNum & "') and (Eligible4EI is null or Eligible4EI = '' or Eligible4EI = 'A_E' or  Eligible4EI = 'ANE')"
    '120     End If
    ' ----------------------------------------------------Previous code


80      If EI.PreviousSample <> 1 Then    'IF no previous sample SET Eligible4EI status = "" <blank>
90          sql = "UPDATE PatientDetails SET Eligible4EI = '' WHERE (labnumber = '" & tLabNum & "') and (Eligible4EI is null or Eligible4EI = '' or Eligible4EI = 'A_E' or  Eligible4EI = 'ANE')"
100     Else
110         If EI.PreviousSample = 1 And EI.PreviousGroupAgreement = 1 And EI.CurrentNegativeAB = 1 And EI.PreviousNegativeAB = 1 _
               And EI.AdverseReactions = 0 And EI.PreviousSampleEligible = 1 And EI.ResultAbnormalFlags = 1 Then    'All 7 Green ticks then mark sample as ELIGIBLE
    'Set as A_E (Automatically Eligible for Electronic Issue) where it was not Forcibly set previously
120             sql = "UPDATE PatientDetails SET Eligible4EI = 'A_E' WHERE (labnumber = '" & tLabNum & "') and (Eligible4EI is null or Eligible4EI = '' or Eligible4EI = 'A_E' or  Eligible4EI = 'ANE')"
130         Else
140             sql = "UPDATE PatientDetails SET Eligible4EI = 'ANE' WHERE (labnumber = '" & tLabNum & "') and (Eligible4EI is null or Eligible4EI = '' or Eligible4EI = 'A_E' or  Eligible4EI = 'ANE')"
150         End If
160     End If

170     CnxnBB(0).Execute sql
180 End If

190 Exit Sub

SaveAutomaticEligibility_Error:

    Dim strES As String
    Dim intEL As Integer

200 intEL = Erl
210 strES = Err.Description
220 LogError "frmxmatch", "SaveAutomaticEligibility", intEL, strES, sql

End Sub


Sub SaveLastUsed()

    Dim LastUsed As String

10  LastUsed = GetSetting("NetAcquire", "Transfusion6", "Lastused")

20  If Val(tLabNum) > Val(LastUsed) Then
30      SaveSetting "NetAcquire", "Transfusion6", "Lastused", tLabNum
40  End If

End Sub
'Zyam commente this code 26-1-24
'Private Sub btnRequestDetail_Click()
'    frmViewProducts.g_SampleID = tLabNum.Text
'    frmViewProducts.Show 1
'
''    Call frmViewProducts.ShowDetail
'End Sub
'Zyam

Private Sub cmbKell_Click()

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub

Private Sub cmbKell_KeyPress(KeyAscii As Integer)

10  KeyAscii = 0

End Sub


Private Sub cmdChangeHistory_Click()

10  frmChangeHistory.txtLabNum = tLabNum
20  frmChangeHistory.cmdSearch.Value = True
30  frmChangeHistory.Show 1


End Sub


Private Sub cmdExternal_Click()

10  txtChart = Trim$(txtChart)
20  tAandE = Trim$(tAandE)

30  If txtChart = "" And tAandE = "" Then
40      iMsg "Notes are tied to Chart or A/E Number." & vbCrLf & _
             "Neither Chart nor A/E Number supplied.", vbExclamation + vbOKOnly
50      If TimedOut Then Unload Me: Exit Sub
60      Exit Sub
70  End If

80  With frmExternal
90      If txtChart <> "" Then
100         .Chart = txtChart
110         .AandE = ""
120     Else
130         .Chart = ""
140         .AandE = tAandE
150     End If
160     .Show 1
170 End With

180 If Not CheckExternalNotes("MRN", txtChart) Then
190     CheckExternalNotes "AandE", tAandE
200 End If

End Sub

Private Function CheckExternalNotes(ByVal ChartOrAandE As String, _
                                    ByVal Criteria As String) _
                                    As Boolean
'Returns True if ExternalNotes present
'ChartOrAandE = "MRN" or "AandE"

    Dim sql As String
    Dim tb As Recordset
    Dim retval As Boolean

10  On Error GoTo CheckExternalNotes_Error

20  retval = False

30  cmdExternal.Caption = "Enter External Notes"
40  cmdExternal.BackColor = vbButtonFace

50  If Trim$(Criteria) <> "" Then
60      sql = "SELECT * FROM ExternalNotes WHERE " & _
              ChartOrAandE & " = '" & Criteria & "'"
70      Set tb = New Recordset
80      RecOpenServerBB 0, tb, sql
90      If Not tb.EOF Then
100         cmdExternal.Caption = "View External Notes"
110         cmdExternal.BackColor = vbYellow
120         retval = True
130     End If
140 End If
150 CheckExternalNotes = retval

160 Exit Function

CheckExternalNotes_Error:

    Dim strES As String
    Dim intEL As Integer

170 intEL = Erl
180 strES = Err.Description
190 LogError "frmxmatch", "CheckExternalNotes", intEL, strES, sql


End Function

Private Sub bgenotype_Click()

10  LoadGenotype tLabNum

20  fgenotype.Show 1
30  labnumberfound = tLabNum

End Sub

Private Sub bHistory_Click()

10  Dept = XMATCH

20  With fpathistory
30      If txtChart <> "" Then
40          .optChart = True
50          .txtName = txtChart
60      Else
70          .optName = True
80          .txtName = txtName
90      End If
100     .cmdSearch.Value = True
110     If Not .NoPrevious Then
120         .Show 1
130     End If
140 End With

150 labnumberfound = tLabNum

End Sub

Private Sub bhold_Click()

10  SaveDetails (True)

End Sub

Private Sub bIssueBatch_Click()

10  If Not SaveDetails(False) Then
20      Exit Sub
30  End If

40  With frmBatchProductIssue
50      .Typenex = tTypenex
60      .SampleID = tLabNum
70      .Show 1
80  End With

90  FillXMHistory

End Sub

Private Sub bPrepare_Click(Index As Integer)

    Dim FaultyGroup As Integer
    Dim s As String
    Dim Rh As String
    Dim pGroup As String
    Dim prh As String

10  If lDontMatch.Visible Then
20      Answer = iMsg("Foward group and Reverse group" & vbCrLf & _
                      "Don't match!" & vbCrLf & _
                      "Do you want to proceed?", vbQuestion + vbYesNo, , vbRed)
30      If TimedOut Then Unload Me: Exit Sub
40      If Answer = vbNo Then
50          Exit Sub
60      End If
70      LogReasonWhy "Prepare(" & tLabNum & "): Forward/Reverse Mis-match. Proceeded with Forward Group Only", "XM"
80  End If

90  If Trim$(lstfg) = "" Then
100     iMsg "Enter Patients Group/Rh!", vbCritical
110     If TimedOut Then Unload Me: Exit Sub
120     Exit Sub
130 End If

140 If InStr(UCase$(lstfg), "POS") Then Rh = "+"
150 If InStr(UCase$(lstfg), "NEG") Then Rh = "-"

160 image2grh pGroup, prh

170 FaultyGroup = False
180 If pGroup <> "" Then
190     If Left$(pGroup & "  ", 2) <> Left$(Trim$(lstfg) & "  ", 2) Then FaultyGroup = True
200     If Rh <> prh Then FaultyGroup = True
210     If prh = "" Then FaultyGroup = True
220     If FaultyGroup Then
230         s = "Difference between" & vbCrLf & _
                "indicated group and" & vbCrLf & _
                "Patients previous group." & vbCrLf & _
                "Continue anyway?"
240         Answer = iMsg(s, vbYesNo + vbQuestion)
250         If TimedOut Then Unload Me: Exit Sub
260         If Answer = vbNo Then
270             Exit Sub
280         End If
290         LogReasonWhy "Prepare(" & tLabNum & "): Indicated/Previous Group Mis-match. Proceeded", "XM"
300     End If
310 End If

320 If lKnownAntibody.Visible Then
330     If Not IsAntibodyHistoryOK() Then Exit Sub
340 End If

350 If Not SaveDetails(False) Then
360     Exit Sub
370 End If

380 If Index = 0 Then    'Manual Prepare
390     With frmXM
400         .ElectronicIssue = False
410         .Show 1
420     End With
430 Else
440     With frmXM
450         .ElectronicIssue = True
460         .Caption = "Electronic Issue"
470         .cmdIssue.Caption = "&E - I"
480         .cXM(0).Visible = False
490         .cXM(1).Visible = False
500         .cXM(2).Visible = False
510         .Picture1(0).Visible = False
520         .Picture1(1).Visible = False
530         .Picture1(2).Visible = False
540         .cmdXM.Visible = False
550         .Show 1
560     End With
570 End If

580 labnumberfound = tLabNum

590 FillXMHistory

600 cmdShowHistory.Caption = "Hide Current"
610 gXmatch.Visible = True

End Sub

Private Sub bPrintDAT_Click()

    Dim Px As Printer
    Dim OriginalPrinter As String
    Dim Y As Integer
    Dim strResult As String
    Dim tb As Recordset
    Dim sql As String
    Dim strAccreditation As String
    Dim TempTb As Recordset
    Dim strName As String
    Dim cColor As ColorConstants
    Dim strResultLine As String

10  On Error GoTo bPrintDAT_Click_Error

20  If Trim$(tLabNum) = "" Then
30      iMsg "Lab Number?", vbQuestion
40      If TimedOut Then Unload Me: Exit Sub
50      Exit Sub
60  End If

70  If Not SaveDetails(True) Then Exit Sub

80  OriginalPrinter = Printer.DeviceName

90  If Not SetFormPrinter() Then Exit Sub

100 strAccreditation = GetOptionSetting("TransfusionAccreditation", "Blood Transfusion at CGH is accredited by INAB to ISO 15189, detailed in scope Registration Number 231MT")

110 With frmRTB

120     PrintHeadingCavan tLabNum

130     PrintTextRTB .rtb, FormatString(" Sample Type: EDTA", 18, , Alignleft), 12, , , , vbRed
140     PrintTextRTB .rtb, FormatString(" Specimen Taken: " & Format(txtSampleDate, "dd/mm/yy hh:mm"), 31, , Alignleft), 12, , , , vbRed
150     PrintTextRTB .rtb, FormatString(" Rec'd Date: " & Format(dtDateRxd, "dd/mm/yy hh:mm") & vbCrLf, 30, , Alignleft) & vbCrLf, 12, , , , vbRed

160     PrintTextRTB .rtb, FormatString("                                                                 ", 35, , AlignCenter), 12, , , , vbBlack
170     PrintTextRTB .rtb, FormatString("DIRECT COOMBS", 14, , AlignCenter) & vbCrLf & vbCrLf, 12, , , 1, vbBlack

180     sql = "SELECT DAT0, DAT1, DAT2, DAT3, DAT4, DAT5, DAT6, DAT7, DAT8, DAT9, DAT10, DAT11 " & _
              "FROM PatientDetails WHERE " & _
              "LabNumber = '" & tLabNum & "'"
190     Set tb = New Recordset
200     RecOpenServerBB 0, tb, sql

210     For Y = 0 To 10 Step 2
220         If tb("DAT" & Format$(Y)) <> 0 Or tb("DAT" & Format$(Y + 1)) <> 0 Then
230             If tb("DAT" & Format$(Y)) <> 0 Then
240                 Printer.ForeColor = vbRed
250                 strResult = " Positive"
260                 cColor = vbRed
270             Else
280                 Printer.ForeColor = vbBlack
290                 strResult = " Negative"
300                 cColor = vbBlack
310             End If
320             strResultLine = Switch(Y = 0, "AHG Poly S    ", Y = 2, "Anti IgG      ", _
                                       Y = 4, "Anti IgA      ", Y = 6, "Anti IgM      ", _
                                       Y = 8, "Anti C3b      ", Y = 10, "Anti C3d      ")

330             strResultLine = strResultLine & strResult
340             PrintTextRTB .rtb, "                                   " & FormatString(strResultLine, 28, , Alignleft) & vbCrLf, 10, , , , cColor
350         End If
360     Next

370     If Trim$(tSampleComment) <> "" Then
380         PrintTextRTB .rtb, FormatString(vbCrLf & vbCrLf & "Comment: " & tSampleComment, 80, , Alignleft) & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf, 10, , , , vbBlack
390     End If

400     PrintTextRTB .rtb, FormatString(vbCrLf & vbCrLf & String$(248, "-"), 248, , AlignCenter) & vbCrLf, 4, , False, , vbRed
410     PrintTextRTB .rtb, FormatString("   " & strAccreditation, 120, , Alignleft) & vbCrLf, 6, , , , vbRed
420     PrintTextRTB .rtb, FormatString(" Report Date:" & Format(Now, "dd/mm/yyyy hh:mm"), 40, , Alignleft), 10, , , , vbRed
430     PrintTextRTB .rtb, FormatString("Issued By " & UserName, 30, , AlignRight), 10, , , , vbRed

440     .rtb.SelPrint Printer.hDC

450     sql = "Select * from PatientDetails where " & _
              "LabNumber = '" & tLabNum & "'"
460     Set TempTb = New Recordset
470     RecOpenServerBB 0, TempTb, sql
480     If TempTb.EOF Then Exit Sub
490     strName = TempTb!Name
    ''''''''''
500     sql = "SELECT * FROM Reports WHERE 0 = 1"
510     Set tb = New Recordset
520     RecOpenServerBB 0, tb, sql
530     tb.AddNew
540     tb!SampleID = tLabNum
550     tb!Name = strName
560     tb!Dept = "DAT Report"
570     tb!Initiator = UserName
580     tb!PrintTime = Now    'PrintTime
590     tb!RepNo = "DAT" & tLabNum & Format(Now, "ddMMyyyyhhmmss")
600     tb!pagenumber = 1
610     tb!Report = .rtb.TextRTF
620     tb!Printer = Printer.DeviceName
630     tb.Update

640 End With

650 For Each Px In Printers
660     If Px.DeviceName = OriginalPrinter Then
670         Set Printer = Px
680         Exit For
690     End If
700 Next

710 Exit Sub

bPrintDAT_Click_Error:

    Dim strES As String
    Dim intEL As Integer

720 intEL = Erl
730 strES = Err.Description
740 LogError "frmxmatch", "bPrintDAT_Click", intEL, strES, sql

End Sub

Private Sub bPrintForm_Click()

    Dim s As String

10  On Error GoTo bPrintForm_Click_Error

20  If lDontMatch.Visible Then
30      Answer = iMsg("Foward group and Reverse group" & vbCrLf & _
                      "Don't match!" & vbCrLf & _
                      "Do you want to report forward group?", vbQuestion + vbYesNo, , vbRed)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60          Exit Sub
70      End If
80      s = iBOX("You must explain why!")
90      If TimedOut Then Unload Me: Exit Sub
100     If Trim$(s) = "" Then
110         iMsg "Operation cancelled!", vbInformation
120         If TimedOut Then Unload Me: Exit Sub
130         Exit Sub
140     End If
150     LogReasonWhy "Print Label(" & tLabNum & "): Forward/Reverse Mis-match. Proceeded with Forward Group Only, " & s, "XM"
160 End If

170 If lstRG.Visible And lstRG = "" Then
180     Answer = iMsg("No Reverse Group!" & vbCrLf & "Proceed?", vbYesNo + vbQuestion, "Incomplete", vbRed)
190     If TimedOut Then Unload Me: Exit Sub
200     If Answer = vbNo Then
210         Exit Sub
220     End If
230     LogReasonWhy "Print Form(" & tLabNum & "): Proceeded with no Reverse Group", "XM"
240 End If

250 If SaveDetails(True) Then

    '260     With frmEligibility
    '270       .Chart = txtChart
    '280       .PatName = txtName
    '290       .DoB = tDoB
    '300       .Show 1
    '310     End With

260     fPrintForm.SampleID = tLabNum
270     fPrintForm.Show 1, Me

280     ShowEligibility

290 End If

300 Exit Sub

bPrintForm_Click_Error:

    Dim strES As String
    Dim intEL As Integer

310 intEL = Erl
320 strES = Err.Description
330 LogError "frmxmatch", "bPrintForm_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10  If cmdSave.Enabled Then
20      Answer = iMsg("Save Details?", vbYesNo + vbQuestion)
30      If TimedOut Then Unload Me: Exit Sub
40      If Answer = vbYes Then
50          If Not SaveDetails(False) Then
60              Exit Sub
70          End If
80      End If
90  End If

100 Unload Me

End Sub

Private Sub btnprint_Click()

    Dim s As String

2   If lKnownAntibody.Visible Then
5     If Not IsAntibodyHistoryOK() Then Exit Sub
8   End If

10  If lDontMatch.Visible Then
20      Answer = iMsg("Foward group and Reverse group" & vbCrLf & _
                      "Don't match!" & vbCrLf & _
                      "Do you want to report forward group?", vbQuestion + vbYesNo, , vbRed)
30      If TimedOut Then Unload Me: Exit Sub
40      If Answer = vbNo Then
50          Exit Sub
60      End If
70      s = iBOX("You must explain why!")
80      If TimedOut Then Unload Me: Exit Sub
90      If Trim$(s) = "" Then
100         iMsg "Operation cancelled!", vbInformation
110         If TimedOut Then Unload Me: Exit Sub
120         Exit Sub
130     End If
140     LogReasonWhy "Print Label(" & tLabNum & "): Forward/Reverse Mis-match. Proceeded with Forward Group Only, " & s, "XM"
150 End If

160 If lstRG.Visible And lstRG = "" Then
170     Answer = iMsg("No Reverse Group!" & vbCrLf & "Proceed?", vbYesNo + vbQuestion, "Incomplete", vbRed)
180     If TimedOut Then Unload Me: Exit Sub
190     If Answer = vbNo Then
200         Exit Sub
210     End If
220     LogReasonWhy "Print Label(" & tLabNum & "): Proceeded with no Reverse Group", "XM"
230 End If

240 If SaveDetails(True) Then
250     frmXMLabel.SampleID = tLabNum
260     frmXMLabel.Show 1, Me
270 End If

280 FillXMHistory

End Sub

Private Sub cmdIDpanels_Click()

10  With frmAnalyserPanels
20      .SampleID = tLabNum
30      .Show 1
40  End With

    'frmIDpanel.Show 1

End Sub

Private Sub cmdIssueToUnknown_Click()

10  If Not SaveDetails(False) Then
20      Exit Sub
30  End If

40  If Trim$(lstfg) <> "" Then
50      bPrepare_Click (0)    'Manual Prepare
60  Else
70      frmXM.Show 1
80  End If

90  labnumberfound = tLabNum

100 FillXMHistory

110 cmdShowHistory.Caption = "Hide Current"
120 gXmatch.Visible = True

End Sub

Private Sub cmdSave_Click()

    Dim sql As String
    'Dim TestRequired As String
    Dim LabelsToPrint As Integer
    Dim intNumOfPrints As Integer
    Dim blnDefaultOrder As Boolean
    Dim ConfirmMRN_MsgBox As New frmCheckMRNmatch
    Dim blnOk2Save As Boolean
    
10  On Error GoTo cmdSave_Click_Error

30  LabelsToPrint = Val(GetOptionSetting("AutoPrintLabels", "2"))
    '40    PrintBarCodesN tLabNum, LabelsToPrint, txtName, txtChart, tDoB, ""
    'If MsgBox("Do you want to print Barcodes?", vbInformation + vbYesNo) = vbYes Then
40      For intNumOfPrints = 1 To LabelsToPrint
50          With frmPrintBarCodeDemo
60              .SampleID = tLabNum
70              .PatientName = txtName
80              .Chart = txtChart
90              .PatientDOB = tDoB
100             .Show 1
110         End With
120     Next
    'End If

'125 With ConfirmMRN_MsgBox
'130     .Show vbModal
'135     blnOk2Save = .retval
'140 End With

'145 Unload ConfirmMRN_MsgBox
'150 Set ConfirmMRN_MsgBox = Nothing
'
'155 If Not blnOk2Save Then
'160     iMsg vbCrLf & "Scan mis-match!" & vbCrLf & vbCrLf & "Patient details and test order not saved!"
'165     Exit Sub
'170 End If

175 If Not SaveDetails(False) Then Exit Sub


280 blnDefaultOrder = True
290 If DateDiff("d", tDoB, txtSampleDate) < 120 Then
300     Answer = iMsg("Is this a Neonatal group?", vbQuestion + vbYesNo)
310     If TimedOut Then Unload Me: Exit Sub
320     If Answer = vbYes Then
330         blnDefaultOrder = False
340     Else
350         blnDefaultOrder = True
360     End If
370 End If

380 If blnDefaultOrder Then    'Default order for over 120 day patients
390     sql = "Insert into BBOrderComms " & _
              "(TestRequired, UnitNumber, SampleID, Programmed) VALUES " & _
              "('FwdRev', '','" & tLabNum & "', 0 )"
400     CnxnBB(0).Execute sql

410     sql = "Insert into BBOrderComms " & _
              "(TestRequired, UnitNumber, SampleID, Programmed) VALUES " & _
              "('Screen Only', '','" & tLabNum & "', 0 )"
420     CnxnBB(0).Execute sql
430 Else    '< than 120 day patients
440     sql = "Insert into BBOrderComms " & _
              "(TestRequired, UnitNumber, SampleID, Programmed) VALUES " & _
              "('Forward Gp', '','" & tLabNum & "', 0 )"
450     CnxnBB(0).Execute sql
460 End If

470 SaveSetting "NetAcquire", "Transfusion6", "LastUsed", tLabNum
    'Zyam 2-1-24
    tLabNum.Text = Val(tLabNum.Text) + 1
    'Zyam
490 'LoadLabNumber

    ClearDetails
   

500 tLabNum.SetFocus

510 Exit Sub

cmdSave_Click_Error:

    Dim strES As String
    Dim intEL As Integer

520 intEL = Erl
530 strES = Err.Description
540 LogError "frmxmatch", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cmdLockDemo_Click()
10  If cmdLockDemo.Caption = "Lock" Then
20      LockDemographics True
30  ElseIf cmdLockDemo.Caption = "Unlock" Then
40      LockDemographics False
50  End If
End Sub

Private Sub cmdOrderAutovue_Click()

10  With frmOrderAutovue
20      .SampleID = tLabNum
30      .Show 1
40  End With

End Sub

Private Sub cmdPhone_Click()

10  With frmPhoneLog
20      .SampleID = Val(tLabNum.Text)
30      If cGP <> "" Then
40          .WardClinGPText = cGP
50          .WardClinGPType = "G"
60      ElseIf cClinician <> "" Then
70          .WardClinGPText = cClinician
80          .WardClinGPType = "C"
90      ElseIf cWard <> "" Then
100         .WardClinGPText = cWard
110         .WardClinGPType = "W"
120     Else
130         iMsg "Either Ward, Clinician or GP must be entered!", vbExclamation
140         If TimedOut Then Unload Me
150         Exit Sub
160     End If
170     .Show 1
180 End With

190 CheckIfPhoned

End Sub

Private Sub CheckIfPhoned()

10  If CheckPhoneLog(tLabNum).SampleID <> 0 Then
20      cmdPhone.BackColor = vbYellow
30      cmdPhone.Caption = "Results Phoned"
40      cmdPhone.ToolTipText = "Results Phoned"
50  Else
60      cmdPhone.BackColor = &H8000000F
70      cmdPhone.Caption = "Phone Results"
80      cmdPhone.ToolTipText = "Phone Results"
90  End If

End Sub

Private Sub cmdSearch_Click()

    Dim f As Form

10  Dept = XMATCH
20  Set f = New frmPatSearch
30  f.From = Me
40  If Trim$(txtChart) <> "" Then
50      f.txtsearch = txtChart
60  Else
70      f.txtsearch = txtName
80  End If
90  f.btncopy.Enabled = True
100 f.btninitiate = True
110 f.bXmatch.Visible = False
120 If f.lblStatus <> "Not found." Then f.Show 1

130 labnumberfound = tLabNum

140 cmdSearch.SetFocus

End Sub

Private Sub bviewwl_Click()

10  If cmdSave.Enabled Then
20      Answer = iMsg("Save Details?", vbYesNo + vbQuestion)
30      If TimedOut Then Unload Me: Exit Sub
40      If Answer = vbYes Then
50          cmdSave.Value = True
60          cmdSave.Enabled = False: bHold.Enabled = False
70      Else
80          Exit Sub
90      End If
100 End If

    'ClearDetails

110 labnumberfound = tLabNum

End Sub

Private Sub ClearDetails()
    Dim n As Integer
    Dim intEI As Integer
    Dim X As Long
    Dim Y As Long

10  For X = 1 To 2
20      gDAT.col = X
30      For Y = 1 To gDAT.Rows - 1
40          gDAT.row = Y
50          Set gDAT.CellPicture = Nothing
60      Next
70  Next

    lKnownAntibody.Visible = False

80  cmbKell = ""
    'Zyam 2-1-24
'90  tLabNum = ""
    'Zyam 2-1-24
100 txtChart = ""
110 tAandE = ""
120 tTypenex = ""
130 txtName = ""
140 txtSurname = ""
150 txtForname = ""
160 tMaiden = ""
170 For n = 0 To 3
180     tAddr(n) = ""
190 Next
200 tDoB = ""
210 tAge = ""
220 tComment = ""
230 lSex = ""
240 cWard = ""
250 cGP = ""
260 cClinician = ""
270 cConditions = ""
280 cProcedure = ""
290 cSpecial = ""
'300 iprevious.Picture = frmMain.ImageList1.ListImages("grPrev").Picture
310 gbitmap = ""

320 optPubPri(0) = 1    'Public
330 optPubPri(1) = 0    'Private


340 lblSuggestRG = ""
350 lstRG = ""

360 lstfg = ""
370 lblsuggestfg.Caption = ""

380 tedd = ""

390 tSampleComment = ""

400 With gXmatch
410     .Rows = 2
420     .AddItem ""
430     .RemoveItem 1
440 End With

450 tident = ""
460 txtSampleDate = ""
470 txtSampleTime = ""

480 lblgrpchecker = ""

490 dtDateRxd = Format(Now, "dd/MM/yyyy")
500 dtTimeRxd = Format(Now, "HH:nn:ss")

'510 lblOrderDAT.BackColor = &H8000000F
520 lblGroupDisagree.Visible = False

'Clear Electronic Issue
  For intEI = 0 To 6
      imgEligible(intEI).Visible = False
  Next

 lblEligible.Visible = False
 lblNotEligible.Visible = False
 lblEligibleTime.Visible = False
 lblForced.Visible = False


End Sub

Private Sub FillXMHistory()

    Dim SamePatient As Boolean
    Dim s As String
    Dim Ps As New Products
    Dim p As Product

10  On Error GoTo FillXMHistory_Error

20  With gXmatch
30      .Visible = False
40      .Rows = 2
50      .AddItem ""
60      .RemoveItem 1
70  End With

80  Ps.LoadLatestBySampleID tLabNum
90  For Each p In Ps
100     If p.ISBT128 <> "" Then
110         s = p.ISBT128
120     Else
130         s = p.PackNumber
140     End If
150     s = s & vbTab & _
            Bar2Group(p.GroupRh & "") & vbTab & _
            Format(p.DateExpiry, "dd/mm/yy HH:mm") & vbTab & _
            p.Screen & vbTab
160     SamePatient = Trim$(p.PatName & "") = Trim$(txtName)
170     Select Case p.PackEvent
            Case "I": s = s & "Issued to " & IIf(SamePatient, "this", "another") & " patient"
180         Case "V": s = s & "EI to " & IIf(SamePatient, "this", "another") & " patient"    'EI to this Patient
190         Case "T": s = s & "Returned to Supplier"
200         Case "K": s = s & "Awaiting Release "
210         Case "S": s = s & "Transfused to " & IIf(SamePatient, "this", "another") & " patient"
220         Case "P": s = s & "Pending for " & IIf(SamePatient, "this", "another") & " patient"
230         Case "D": s = s & "Destroyed"
240         Case "R": s = s & "Restocked"
250         Case "F": s = s & "Transferred with " & IIf(SamePatient, "this", "another") & " patient"
260         Case "X": s = s & "Xmatched to " & IIf(SamePatient, "this", "another") & " patient"
270         Case "Y": s = s & "Removed Pending Transfusion to " & IIf(SamePatient, "this", "another") & " patient"
280     End Select
290     s = s & vbTab
300     If p.crt Then s = s & IIf(p.crtr, "+", "O")
310     s = s & vbTab
320     If p.cco Then s = s & IIf(p.ccor, "+", "O")
330     s = s & vbTab
340     If p.cen Then s = s & IIf(p.cenr, "+", "O")
350     s = s & vbTab
360     s = s & ProductWordingFor(p.Barcode)
370     s = s & vbTab & p.UserName & vbTab & p.RecordDateTime
380     gXmatch.AddItem s
390 Next

400 FillBatchHistory

410 With gXmatch
420     .col = 10
430     .Sort = 9
440     If .Rows > 2 Then
450         .RemoveItem 1
460     End If
470     .Visible = True
480 End With

490 Exit Sub

FillXMHistory_Error:

    Dim strES As String
    Dim intEL As Integer

500 intEL = Erl
510 strES = Err.Description
520 LogError "frmxmatch", "FillXMHistory", intEL, strES
530 gXmatch.Visible = True

End Sub

Private Sub FillBatchHistory()

    Dim s As String
    Dim BP As BatchProduct
    Dim BPs As New BatchProducts

10  On Error GoTo FillBatchHistory_Error

20  BPs.LoadSampleIDNoAudit tLabNum
30  For Each BP In BPs

40      s = BP.BatchNumber & vbTab & _
            BP.UnitGroup & vbTab & _
            Format$(BP.DateExpiry, "dd/mm/yy") & vbTab & vbTab & _
            gEVENTCODES(BP.EventCode).Text & _
            vbTab & vbTab & vbTab & vbTab & _
            BP.Product & vbTab & _
            BP.UserName & vbTab & _
            Format(BP.RecordDateTime, "dd/mm/yyyy hh:mm:ss") & vbTab
50      If Format$(BP.EventStart, "dd/MM/yyyy") <> "01/01/1900" Then
60          s = s & Format$(BP.EventStart, "dd/MM/yy HH:nn:ss")
70      End If
80      s = s & vbTab
90      If Format$(BP.EventEnd, "dd/MM/yyyy") <> "01/01/1900" Then
100         s = s & Format$(BP.EventEnd, "dd/MM/yy HH:nn:ss")
110     End If
120     s = s & vbTab & BP.Identifier
130     gXmatch.AddItem s

140 Next

150 Exit Sub

FillBatchHistory_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmxmatch", "FillBatchHistory", intEL, strES

End Sub

Private Sub cClinician_LostFocus()

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo cClinician_LostFocus_Error

20  cmdSave.Enabled = True
30  bHold.Enabled = True

40  sql = "Select * from Clinicians where " & _
          "Code = '" & AddTicks(cClinician) & "' " & _
          "or Text = '" & AddTicks(cClinician) & "'"
50  Set tb = New Recordset
60  RecOpenServer 0, tb, sql
70  If Not tb.EOF Then
80      cClinician = tb!Text & ""
90  Else
100     cClinician = ""
110 End If

120 Exit Sub

cClinician_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmxmatch", "cClinician_LostFocus", intEL, strES, sql


End Sub


Private Sub cConditions_LostFocus()

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo cConditions_LostFocus_Error

20  cmdSave.Enabled = True
30  bHold.Enabled = True

40  sql = "Select * from Lists where " & _
          "ListType = 'X' " & _
          "and Code = '" & Trim$(UCase$(cConditions)) & "'"
50  Set tb = New Recordset
60  RecOpenServer 0, tb, sql
70  If Not tb.EOF Then
80      cConditions = tb!Text & ""
90  End If

100 Exit Sub

cConditions_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmxmatch", "cConditions_LostFocus", intEL, strES, sql


End Sub


Private Sub cGP_LostFocus()

    Dim tb As Recordset
    Dim sql As String
    Dim strGP As String

10  On Error GoTo cGP_LostFocus_Error

20  cmdSave.Enabled = True
30  bHold.Enabled = True

40  strGP = AddTicks(cGP)

50  sql = "Select * from GPs where " & _
          "Code = '" & strGP & "' " & _
          "or [Text] = '" & strGP & "'"
60  Set tb = New Recordset
70  RecOpenServer 0, tb, sql
80  If Not tb.EOF Then
90      cGP = tb!Text & ""
100 Else
110     cGP = ""
120 End If

130 Exit Sub

cGP_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmxmatch", "cGP_LostFocus", intEL, strES, sql


End Sub


Private Sub cmdLab_Click()

10  If Trim$(txtChart) = "" Then
20      iMsg "Must Have Chart Number to View Lab Results!", vbInformation
30      If TimedOut Then Unload Me: Exit Sub
40      Exit Sub
50  End If

60  frmViewGeneral.Show 1

End Sub

Private Sub cmdShowHistory_Click()

10  If cmdShowHistory.Caption = "Show Current" Then
20      FillXMHistory
30      cmdShowHistory.Caption = "Hide Current"
40      gXmatch.Visible = True
50  Else
60      cmdShowHistory.Caption = "Show Current"
70      gXmatch.Visible = False
80  End If

End Sub

Private Sub cmdSuggestUnits_Click()

    Dim NumberOfPacks As Integer

10  If Trim$(lstfg) = "" Then
20      iMsg "Patients Group must be specified.", vbOKOnly
30      If TimedOut Then Unload Me
40      Exit Sub
50  End If

60  NumberOfPacks = Val(iBOX("How many Packs?", , "1"))
    'If NumberOfPacks < 1 Or NumberOfPacks > 10 Then Exit Sub

70  With frmSuggestFromStock
80      .PatName = txtName
90      .DoB = tDoB
100     .Age = tAge
110     .Sex = lSex
120     .SampleID = tLabNum
130     .Group = lstfg
140     .Kell = cmbKell
150     .NumberOfPacks = NumberOfPacks
160     .Show 1
170 End With

End Sub



Private Sub cmdUpDown_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim blnIncrement As Boolean

blnIncrement = False


0    If Index = 0 Then
30        UpDownDirection = 1
          If Val(Mid(tLabNum, 8, 5)) < 99999 Then blnIncrement = True
40    Else
50        UpDownDirection = -1
          If Val(tLabNum) > 1 Then blnIncrement = True
60    End If

If blnIncrement Then
      If Len(tLabNum) > 12 Then
         tLabNum = Val(tLabNum) + Val(UpDownDirection)
'        tLabNum = Left(tLabNum, 7) & Format(Val(Mid(tLabNum, 8, 5)) + UpDownDirection, "") & Mid(tLabNum, 13, 1)
      Else
'80      tLabNum = Left(tLabNum, 7) & Format(Val(Mid(tLabNum, 8, 5)) + UpDownDirection, "")
        tLabNum = Val(tLabNum) + Val(UpDownDirection)
      End If
      
End If

End Sub

Private Sub cmdUpDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

UpDownDirection = 0

pBar = 0

LoadLabNumber

End Sub

Private Sub cmdValidate_Click()
10  If cmdValidate.Caption = "Confirm Details" Then
20      If ValidForConfirmation Then
30          SecondUserName = UserName
40          SaveDetails True
50          frmConfirmDetails.Show 1
60          LockUnlockConfirmedMallow
70      End If
80  ElseIf cmdValidate.Caption = "UnConfirm" Then
90      If AuthenticateUser(Managers) Then
100         frmConfirmDetails.Show 1
110         LockUnlockConfirmedMallow
120     Else
130         iMsg "Invalid password", vbInformation
140         If TimedOut Then Unload Me: Exit Sub
150     End If
160 End If
End Sub

Private Sub cmdViewEligibility_Click()

10  If Not IsDate(txtSampleDate) Then
20      iMsg "Enter Sample Date", vbCritical, , vbRed
30      If TimedOut Then Unload Me
40      Exit Sub
50  End If

60  If Not IsDate(txtSampleTime) Then
70      iMsg "Enter Sample Time", vbCritical, , vbRed
80      If TimedOut Then Unload Me
90      Exit Sub
100 End If

110 With frmEligibility
120     .SampleDate = txtSampleDate & " " & txtSampleTime
130     .SampleID = tLabNum
140     .Chart = txtChart
150     .PatName = txtName
160     .DoB = tDoB
170     .Show 1
180 End With

190 ShowEligibility

End Sub


Private Sub cmdViewReports_Click()
    Dim f As Form

10  Set f = New frmReportViewer

20  f.Dept = "XMatch"
30  f.SampleID = tLabNum
40  f.Show 1

50  Set f = Nothing
End Sub



Private Sub dtDateRxd_CloseUp()

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub


Private Sub dtTimeRxd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

10  If KeyCode = vbKeyF2 Then
20      cmdSave.Value = True
30  End If

End Sub


Private Sub gDAT_Click()

10  If gDAT.MouseRow = 0 Or gDAT.MouseCol = 0 Then Exit Sub

20  gDAT.col = gDAT.MouseCol
30  gDAT.row = gDAT.MouseRow

40  If gDAT.CellPicture <> imgSquareCross.Picture Then
50      Set gDAT.CellPicture = imgSquareCross.Picture
60  Else
70      Set gDAT.CellPicture = Nothing    ''imgSquare.Picture
80  End If
90  gDAT.col = IIf(gDAT.col = 1, 2, 1)
100 Set gDAT.CellPicture = Nothing

110 gDAT.CellPictureAlignment = flexAlignCenterCenter
120 cmdSave.Enabled = True
130 bHold.Enabled = True

End Sub

Private Sub imgUseTime_Click()

'10  If imgUseTime.Picture = imgSquareCross.Picture Then
'20      Set imgUseTime.Picture = imgSquareTick.Picture
'30      dtTimeRxd.Enabled = True
'40      dtTimeRxd = Format(Now, "HH:nn:ss")
'50  Else
'60      Set imgUseTime.Picture = imgSquareCross.Picture
'70      dtTimeRxd.Enabled = False
'80  End If

End Sub

Private Sub iprevious_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10  If txtChart = "" Then Exit Sub
20  If lblGroupDisagree.Visible = True Then Exit Sub

    Dim tb As Recordset
    Dim sql As String
    Dim s As String
30  Set tb = New Recordset
40  sql = "Select prevgroup, previousrh From PatientDetails Where patnum = '" & txtChart & "'"
50  RecOpenClientBB 0, tb, sql
60  If (Not IsNull(tb!PrevGroup)) And tb!PrevGroup <> "" _
       And (Not IsNull(tb!previousrh)) And tb!previousrh <> "" Then
70      Select Case tb!previousrh
            Case "+"
80              s = "Pos"
90          Case "-"
100             s = "Neg"
110     End Select
120     s = "gr" & tb!PrevGroup & s
'130     iprevious.Picture = frmMain.ImageList1.ListImages(s).Picture

140 End If
150 tb.Close
End Sub

Private Sub iprevious_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'10  iprevious.Picture = frmMain.ImageList1.ListImages("grPrev").Picture
End Sub


'Private Sub lblOrderDAT_Click()
'
'10  If lblOrderDAT.BackColor = vbGreen Then
'20      lblOrderDAT.BackColor = vbButtonFace
'30  Else
'40      lblOrderDAT.BackColor = vbGreen
'50  End If
'
'60  cmdSave.Enabled = True
'
'End Sub

Private Sub lstfg_LostFocus()

10  lstfg.Enabled = True

End Sub

Private Sub cProcedure_LostFocus()

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo cProcedure_LostFocus_Error

20  cmdSave.Enabled = True
30  bHold.Enabled = True

40  sql = "Select * from Lists where " & _
          "ListType = 'P' " & _
          "and Code = '" & Trim$(UCase$(cProcedure)) & "'"
50  Set tb = New Recordset
60  RecOpenClientBB 0, tb, sql
70  If Not tb.EOF Then
80      cProcedure = tb!Text & ""
90  End If

100 Exit Sub

cProcedure_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmxmatch", "cProcedure_LostFocus", intEL, strES, sql


End Sub


Private Sub cSpecial_LostFocus()

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo cSpecial_LostFocus_Error

20  cmdSave.Enabled = True
30  bHold.Enabled = True

40  sql = "Select * from Lists where " & _
          "ListType = 'S' " & _
          "and Code = '" & Trim$(UCase$(cSpecial)) & "'"
50  Set tb = New Recordset
60  RecOpenClient 0, tb, sql
70  If Not tb.EOF Then
80      cSpecial = tb!Text & ""
90  End If


100 Exit Sub

cSpecial_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmxmatch", "cSpecial_LostFocus", intEL, strES, sql


End Sub


Private Sub cWard_LostFocus()

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo cWard_LostFocus_Error

20  cmdSave.Enabled = True
30  bHold.Enabled = True

40  sql = "Select * from Wards where " & _
          "Code = '" & AddTicks(cWard) & "' " & _
          "or Text = '" & AddTicks(cWard) & "'"

50  Set tb = New Recordset
60  RecOpenServer 0, tb, sql
70  If tb.EOF Then
80      cWard = ""
90  Else
100     cWard = tb!Text & ""
110 End If

120 Exit Sub

cWard_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmxmatch", "cWard_LostFocus", intEL, strES, sql


End Sub


Private Sub Form_Activate()

    Static Loaded As Boolean

10  If Not Loaded Then
20      tLabNum.SetFocus
30      Loaded = True
40  End If

50  Loading = False
    timFetchSampleID.Enabled = True

End Sub

Private Sub NameLostFocus(ByRef strName As String, _
                          ByRef strSex As String)

    Dim ForeName As String
    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo NameLostFocus_Error

20  strName = Replace(strName, ",", "")

30  strName = Initial2Upper(strName)

40  ForeName = ParseForeName(strName)
50  If ForeName = "" Then
60      strSex = ""
70      Exit Sub
80  End If

90  sql = "Select * from SexNames where " & _
          "Name = '" & AddTicks(ForeName) & "'"
100 Set tb = New Recordset
110 RecOpenServer 0, tb, sql
120 If tb.EOF Then
130     If strSex <> "" Then
140         tb.AddNew
150         tb!Name = ForeName
160         tb!Sex = UCase$(Left$(strSex, 1))
170         tb.Update
180     End If
190 Else
200     Select Case UCase$(tb!Sex & "")
            Case "M": strSex = "Male"
210         Case "F": strSex = "Female"
220     End Select
230 End If

240 Exit Sub

NameLostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

250 intEL = Erl
260 strES = Err.Description
270 LogError "frmxmatch", "NameLostFocus", intEL, strES, sql


End Sub

Private Sub FillLists()

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo FillLists_Error

20  cWard.Clear
30  cGP.Clear
40  cClinician.Clear
50  cConditions.Clear
60  cProcedure.Clear
70  cSpecial.Clear

80  sql = "SELECT DISTINCT Text, MIN(ListOrder) L FROM Wards WHERE " & _
          "InUse = 1 " & _
          "GROUP BY Text " & _
          "ORDER BY L"
90  Set tb = New Recordset
100 RecOpenServer 0, tb, sql
110 Do While Not tb.EOF
120     cWard.AddItem tb!Text & ""
130     tb.MoveNext
140 Loop

150 sql = "SELECT DISTINCT Text, MIN(ListOrder) L FROM Clinicians " & _
          "WHERE InUse = 1 " & _
          "GROUP BY Text " & _
          "ORDER BY L"
160 Set tb = New Recordset
170 RecOpenServer 0, tb, sql
180 Do While Not tb.EOF
190     cClinician.AddItem tb!Text & ""
200     tb.MoveNext
210 Loop

220 sql = "SELECT DISTINCT Text, MIN(ListOrder) L FROM GPs " & _
          "WHERE InUse = 1 " & _
          "GROUP BY Text " & _
          "ORDER BY L"
230 Set tb = New Recordset
240 RecOpenServer 0, tb, sql
250 Do While Not tb.EOF
260     cGP.AddItem tb!Text & ""
270     tb.MoveNext
280 Loop

290 sql = "SELECT Text, ListType FROM Lists WHERE " & _
          "ListType = 'X' " & _
          "OR ListType = 'P' " & _
          "OR ListType = 'S' " & _
          "ORDER BY ListOrder"
300 Set tb = New Recordset
310 RecOpenServerBB 0, tb, sql
320 Do While Not tb.EOF
330     If tb!ListType = "X" Then
340         cConditions.AddItem tb!Text & ""
350     ElseIf tb!ListType = "P" Then
360         cProcedure.AddItem tb!Text & ""
370     ElseIf tb!ListType = "S" Then
380         cSpecial.AddItem tb!Text & ""
390     End If
400     tb.MoveNext
410 Loop

420 Exit Sub

FillLists_Error:

    Dim strES As String
    Dim intEL As Integer

430 intEL = Erl
440 strES = Err.Description
450 LogError "frmxmatch", "FillLists", intEL, strES, sql

End Sub


Private Sub Form_Load()

    Dim Added As Boolean
    Dim X As Long
    Dim Y As Long

10  On Error GoTo Form_Load_Error

20  udLabNum.max = 999999
30  dtDateRxd = Format(Now, "dd/mm/yyyy")

40  Added = False

50  lstfg.Font.Size = 8
60  lstfg.Width = 1035

70  gDAT.AddItem "AHG Poly-S"
80  gDAT.AddItem "Anti IgG"
90  gDAT.AddItem "Anti IgA"
100 gDAT.AddItem "Anti IgM"
110 gDAT.AddItem "Anti C3b,C3d"
120 gDAT.Height = 1455
130 gDAT.RemoveItem 1
140 For X = 1 To 2
150     gDAT.col = X
160     For Y = 1 To gDAT.Rows - 1
170         gDAT.row = Y
180         gDAT.CellPictureAlignment = flexAlignCenterCenter
190     Next
200 Next

210 cmbKell.Clear
220 cmbKell.AddItem ""
230 cmbKell.AddItem "K-"
240 cmbKell.AddItem "K+"
250 lblChartTitle.Caption = "Chart"
260 lstfg.Font.Size = 12
270 lstfg.Width = 1905
280 cmdShowHistory.Caption = "Hide Current"
290 gXmatch.Visible = True
300 lblSuggestRG.Visible = False
310 lstRG.Visible = False
320 lblNOPAS.Visible = False
330 txtNOPAS.Visible = False
340 cmdOrderAutovue.Visible = True

350 lblChartNumber.Caption = HospName(0) & " Chart #"

360 Dept = XMATCH
370 FillLists
380 Fill_FG

390 SuggestLabNum
400 LoadLabNumber
410 cmdSave.Enabled = False
420 bHold.Enabled = False

430 Exit Sub

Form_Load_Error:

    Dim strES As String
    Dim intEL As Integer

440 intEL = Erl
450 strES = Err.Description
460 LogError "frmxmatch", "Form_Load", intEL, strES

End Sub
Private Sub Fill_FG()

    Dim n As Integer
    Dim s As String

10  With lstfg
20      For n = 0 To 12
30          s = Choose(n + 1, "", _
                       "O Neg", _
                       "O Pos", _
                       "A Neg", _
                       "A Pos", _
                       "B Neg", _
                       "B Pos", _
                       "AB Neg", _
                       "AB Pos", _
                       "O Dvi", _
                       "A Dvi", _
                       "B Dvi", _
                       "AB Dvi")
40          .AddItem s, n
50      Next
60      .AddItem "Control ?", 13
70      .AddItem "Error", 14
80  End With

End Sub

Private Sub gXmatch_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

    Dim d1 As String
    Dim d2 As String

10  If Not IsDate(gXmatch.TextMatrix(Row1, 10)) Then
20      Cmp = 0
30      Exit Sub
40  End If

50  If Not IsDate(gXmatch.TextMatrix(Row2, 10)) Then
60      Cmp = 0
70      Exit Sub
80  End If

90  d1 = Format(gXmatch.TextMatrix(Row1, 10), "dd/mmm/yyyy hh:mm:ss")
100 d2 = Format(gXmatch.TextMatrix(Row2, 10), "dd/mmm/yyyy hh:mm:ss")

110 Cmp = Sgn(DateDiff("D", d1, d2))

End Sub


Private Sub iprevious_Click()

10  If lblGroupDisagree.Visible = False Then Exit Sub

    Dim tb As Recordset
    Dim sql As String
    Dim s As String
    Dim strPass As String
    Dim f As Form
    Dim NewGRh As String
    Dim PW As String

20  On Error GoTo iprevious_Click_Error

30  s = "Amend Historical Group?" & vbCrLf & "(Only for this Lab Number)"
40  Answer = iMsg(s, vbQuestion + vbYesNo)
50  If TimedOut Then Unload Me: Exit Sub
60  If Answer = vbYes Then
70      sql = "Select Password from Users where " & _
              "Name = '" & AddTicks(UserName) & "' " & _
              "COLLATE SQL_Latin1_General_CP1_CS_AS"

80      Set tb = New Recordset
90      RecOpenServer 0, tb, sql
100     If Not tb.EOF Then
110         strPass = tb!Password & ""
120         PW = UCase$(iBOX("Password", , , True))
130         If TimedOut Then Unload Me: Exit Sub
140         If PW = UCase$(strPass) Then
150             Set f = frmSelectGroup
160             f.Show 1
170             NewGRh = f.GRh
180             Set f = Nothing
190             If NewGRh = "" Then Exit Sub

200             sql = "Insert into IncidentLog " & _
                      "(DateTime, Incident, Technician) VALUES " & _
                      "('" & Format$(Now, "dd/mmm/yyyy") & "', " & _
                      "'Previous Group changed to " & NewGRh & "', " & _
                      "'" & UserName & "')"
210             CnxnBB(0).Execute sql

220             s = "gr" & NewGRh
230             Select Case UCase$(NewGRh)
                    Case "ONEG": gbitmap = "o-"
240                 Case "OPOS": gbitmap = "o+"
250                 Case "ANEG": gbitmap = "a-"
260                 Case "APOS": gbitmap = "a+"
270                 Case "BNEG": gbitmap = "b-"
280                 Case "BPOS": gbitmap = "b+"
290                 Case "ABNEG": gbitmap = "ab-"
300                 Case "ABPOS": gbitmap = "ab+"

310             End Select

320             'iprevious.Picture = frmMain.ImageList1.ListImages(s).Picture
330             cmdSave.Enabled = True: bHold.Enabled = True
340         End If
350     End If
360 End If

370 Exit Sub

iprevious_Click_Error:

    Dim strES As String
    Dim intEL As Integer

380 intEL = Erl
390 strES = Err.Description
400 LogError "frmxmatch", "iprevious_Click", intEL, strES, sql


End Sub

Private Sub lblChartNumber_Click()
'
'10    With lblChartNumber
'20      .BackColor = &H8000000F
'30      .ForeColor = vbBlack
'40      If .Caption = "Monaghan Chart #" Then
'50        .Caption = "Cavan Chart #"
'60      ElseIf .Caption = "Cavan Chart #" Then
'70        .Caption = "Monaghan Chart #"
'80        If UCase$(HospName(0)) = "CAVAN" Then
'90          .BackColor = vbRed
'100         .ForeColor = vbYellow
'110       End If
'120     End If
'130   End With
'
'140   If Trim$(txtChart) <> "" Then
'150     txtChart_LostFocus
'160     cmdLab.Visible = True
'170   Else
'180     cmdLab.Visible = False
'190   End If

End Sub

Private Sub lblgrpchecker_Click()

10  cmdSave.Enabled = True
20  bHold.Enabled = True

End Sub

Private Sub lInfo_Click()

10  lInfo.Visible = False

End Sub

Private Sub lsex_Click()

'If Not PatientDetailsConfirmed(tLabNum) Then

10  Select Case Left$(lSex, 1)
        Case "M": lSex = "Female"
20      Case "F": lSex = "Unknown"
30      Case "U": lSex = ""
40      Case Else: lSex = "Male"
50  End Select

60  cmdSave.Enabled = True: bHold.Enabled = True

    'End If

End Sub

Private Sub lstFG_Click()

    Dim PrevRH As String
    Dim PrevGroup As String

10  cmdSave.Enabled = True: bHold.Enabled = True
20  lblGroupDisagree.Visible = False

30  If Trim$(txtChart) = "" Then Exit Sub

40  PrevGroup = GroupKnown(txtChart, tLabNum, txtName)

50  If PrevGroup <> "" Then
60      If UCase$(Trim$(PrevGroup)) <> UCase$(Trim$(lstfg)) Then
70          lblGroupDisagree.Visible = True
80          If UCase(iBOX("Historical Grouping Disagreement !" & vbCrLf & vbCrLf & "Enter Password", , , True)) <> UCase(UserPasswordForCode(UserCode)) Then
90              lstfg = ""
100             Exit Sub
110         Else
120             LogReasonWhy "Historical Grouping Disagreement - warning issued. (Lab number " & tLabNum & ")", "HGD"
130         End If
140     End If
150     PrevRH = ""
160     If InStr(UCase$(PrevGroup), "POS") Then PrevRH = "+"
170     If InStr(UCase$(PrevGroup), "NEG") Then PrevRH = "-"
180     If InStr(PrevGroup, "-") Then PrevRH = "-"
190     If InStr(PrevGroup, "+") Then PrevRH = "+"
200     PrevGroup = Trim$(Left$(PrevGroup, 2))
210     grh2image PrevGroup, PrevRH
220 End If

End Sub

Private Function GroupHistoryOK() As Boolean

    Dim sql As String
    Dim tb As Recordset
    Dim Group As String

10  On Error GoTo GroupHistoryOK_Error

20  GroupHistoryOK = True
30  If Trim$(txtChart) = "" Then
40      cmdLab.Visible = False
50      Exit Function
60  End If

70  cmdLab.Visible = True

80  sql = "select fGroup from PatientDetails where " & _
          "PatNum = '" & txtChart & "' "
    '"and Name = '" & AddTicks(txtName) & "'"
    'Line above removed in accordance with Ken's instructions 9 Aug 2013
90  Set tb = New Recordset
100 RecOpenServerBB 0, tb, sql

110 Group = ""
120 Do While Not tb.EOF
130     If Trim$(tb!fGroup & "") <> "" Then
140         If Group = "" Then
150             Group = tb!fGroup
160         Else
170             If UCase$(Trim$(Group)) <> UCase$(Trim$(tb!fGroup)) Then
180                 GroupHistoryOK = False
190                 Exit Function
200             End If
210         End If
220     End If
230     tb.MoveNext
240 Loop

250 Exit Function

GroupHistoryOK_Error:

    Dim strES As String
    Dim intEL As Integer

260 intEL = Erl
270 strES = Err.Description
280 LogError "frmxmatch", "GroupHistoryOK", intEL, strES, sql


End Function

Private Sub lstfg_KeyPress(KeyAscii As Integer)

10  KeyAscii = 0

End Sub


Private Sub lstrg_Change()

10  lDontMatch.Visible = False

20  If Trim$(lstfg) <> "" And Trim$(lstRG) <> "" Then
30      If Left$(lstfg, 2) <> Left$(lstRG & " ", 2) Then
40          lDontMatch.Visible = True
50      End If
60  End If

End Sub

Private Sub lstRG_Click()

    Dim FaultyGroup As Integer
    Dim Rh As String
    Dim s As String
    Dim pGroup As String
    Dim prh As String

10  If InStr(UCase$(lstfg), "POS") Then Rh = "+"
20  If InStr(UCase$(lstfg), "NEG") Then Rh = "-"

    'image2grh iprevious.Picture, pgroup, prh

30  FaultyGroup = False
40  If Trim$(lstRG) <> "" And pGroup <> "" Then
50      If Left$(pGroup & "  ", 2) <> Left$(Trim$(lstRG) & "  ", 2) Then FaultyGroup = True
60      If Rh <> prh Then FaultyGroup = True
70      If FaultyGroup Then
80          s = "Difference between" & vbCrLf & _
                "indicated group and" & vbCrLf & _
                "Patients previous group."
90          iMsg s, vbExclamation
100         If TimedOut Then Unload Me: Exit Sub
110     End If
120 End If

130 lDontMatch.Visible = False
140 If Trim$(lstfg) <> "" And Trim$(lstRG) <> "" Then
150     If Left$(lstfg, 2) <> Left$(lstRG & " ", 2) Then
160         lDontMatch.Visible = True
170     End If
180 End If

190 If lstRG <> lblSuggestRG And lblSuggestRG <> "" Then
200     lDontMatch.Visible = True
210 End If

220 cmdSave.Enabled = True: bHold.Enabled = True

End Sub


Private Sub lstrg_KeyPress(KeyAscii As Integer)

10  KeyAscii = 0

End Sub


Private Sub optPubPri_Click(Index As Integer)
10  cmdSave.Enabled = True: bHold.Enabled = True
End Sub

Private Sub tAandE_KeyPress(KeyAscii As Integer)

'If PatientDetailsConfirmed(tLabNum) Then
'  KeyAscii = 0
'Else
10  cmdSave.Enabled = True: bHold.Enabled = True
    'End If

End Sub

Private Sub tAddr_KeyPress(Index As Integer, KeyAscii As Integer)

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub

Private Sub tAge_KeyPress(KeyAscii As Integer)

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub

Private Sub tComment_KeyPress(KeyAscii As Integer)

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub

Private Sub tDoB_KeyPress(KeyAscii As Integer)

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub

Private Sub tedd_KeyPress(KeyAscii As Integer)

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub


Private Sub tident_KeyPress(KeyAscii As Integer)

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub

Public Sub timFetchSampleID_Timer()
    If g_SampleID = "" Then
'        timFetchSampleID.Enabled = False
        Exit Sub
    End If
    Call tLabNum_GotFocus
    DoEvents
    DoEvents
    tLabNum.Text = g_SampleID
    Call tLabNum_KeyPress(13)
    DoEvents
    DoEvents
    g_SampleID = ""
'    timFetchSampleID.Enabled = False
End Sub

Private Sub tLabNum_KeyPress(KeyAscii As Integer)

10  If KeyAscii = 32 Then    'if SPACE character pressed
20      KeyAscii = 0
30  End If
    '+++ Junaid 08-07-2022
    If KeyAscii = 13 Then
        Call tlabnum_LostFocus
    End If
    '--- Junaid 08-07-2022
End Sub

Private Sub tMaiden_KeyPress(KeyAscii As Integer)

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

'If PatientDetailsConfirmed(tLabNum) Then
'  KeyAscii = 0
'Else
10  cmdSave.Enabled = True: bHold.Enabled = True
    'End If

End Sub

Private Sub tTypenex_KeyPress(KeyAscii As Integer)

'If PatientDetailsConfirmed(tLabNum) Then
'  KeyAscii = 0
'Else
10  KeyAscii = MaskInput(KeyAscii, tTypenex, "XXX ####")
20  cmdSave.Enabled = True: bHold.Enabled = True
    'End If

End Sub

Private Sub txtChart_KeyPress(KeyAscii As Integer)

'If PatientDetailsConfirmed(tLabNum) Then
'  KeyAscii = 0
'Else
10  cmdSave.Enabled = True: bHold.Enabled = True
    'End If

End Sub

Private Sub txtNOPAS_KeyPress(KeyAscii As Integer)

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub

Private Sub txtSampleTime_KeyPress(KeyAscii As Integer)

10  Select Case Chr(KeyAscii)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ":"
20      Case Else: KeyAscii = 0
30  End Select

40  cmdSave.Enabled = True: bHold.Enabled = True

End Sub

Private Sub txtSampleTime_LostFocus()

10  If Len(txtSampleTime) = 4 Then
20      txtSampleTime = Left$(txtSampleTime, 2) & ":" & Right$(txtSampleTime, 2)
30  End If
40  If Not IsDate(txtSampleTime) Then
50      txtSampleTime = ""
60  End If

End Sub

Private Function SaveDetails(ByVal Hold As Integer) As Boolean

    Dim n As Integer
    Dim s As String
    Dim ds As Recordset
    Dim sql As String
    Dim pGroup As String
    Dim prh As String
    Dim TimeNow As String
    Dim HospitalName As String
    Dim DateTimeRxd As String
    Dim retval As Boolean

10  On Error GoTo SaveDetails_Error

20  SaveDetails = False

30  lblGroupDisagree.Visible = False
40  If Not GroupHistoryOK() Then
50      lblGroupDisagree.Visible = True
60      If UCase(iBOX("Historical Grouping Disagreement !" & vbCrLf & vbCrLf & "Enter Password", , , True)) <> UCase(UserPasswordForCode(UserCode)) Then
70          Exit Function
80      Else
90          LogReasonWhy "Historical Grouping Disagreement - warning issued. (Lab number " & tLabNum & ")", "HGD"
100     End If
110 End If

120 If lDontMatch.Visible Then
130     Answer = iMsg("Foward group and Reverse group" & vbCrLf & _
                      "Don't match!" & vbCrLf & _
                      "Do you want to report forward group?", vbQuestion + vbYesNo, , vbRed)
140     If TimedOut Then Unload Me: Exit Function
150     If Answer = vbNo Then
160         Exit Function
170     End If
180     s = iBOX("You must explain why!")
190     If TimedOut Then Unload Me: Exit Function
200     If Trim$(s) = "" Then
210         iMsg "Operation cancelled!", vbInformation
220         If TimedOut Then Unload Me: Exit Function
230         Exit Function
240     End If
250     LogReasonWhy "Save Details(" & tLabNum & "): Forward/Reverse Mis-match. Proceeded with Forward Group Only, " & s, "XM"
260 End If

270 SaveSetting "NetAcquire", "Transfusion6", "LastUsed", tLabNum

280 If Trim$(tLabNum) = "" Then
290     iMsg "Specify Lab Number!", vbCritical
300     If TimedOut Then Unload Me: Exit Function
310     Exit Function
320 End If

330 If Trim$(txtChart) = "" And Trim$(txtName) = "" Then
340     iMsg "Require either Patients Name" & vbCrLf & "or number.", vbCritical
350     If TimedOut Then Unload Me: Exit Function
360     Exit Function
370 End If

380 If Not IsDate(txtSampleDate) Then
390     iMsg "Sample Date must be entered.", vbExclamation
400     If TimedOut Then Unload Me: Exit Function
410     Exit Function
420 End If

430 If Not IsDate(txtSampleTime) Then
440     iMsg "Sample Time must be entered.", vbExclamation
450     If TimedOut Then Unload Me: Exit Function
460     Exit Function
470 End If

480 If Trim$(lSex) = "" Then
490     iMsg "Sex must be entered!", vbExclamation
500     If TimedOut Then Unload Me: Exit Function
510     Exit Function
520 End If

530 If Trim$(tDoB) = "" Then
540     Answer = iMsg("No Date of Birth entered!" & vbCrLf & "Proceed?", vbQuestion + vbYesNo)
550     If TimedOut Then Unload Me: Exit Function
560     If Answer = vbNo Then
570         Exit Function
580     Else
590         LogReasonWhy "Save Details (" & tLabNum & "): Proceeded with no Date of Birth", "XM"
600     End If
610 End If

620 TimeNow = Format$(Now, "dd/mm/yyyy hh:mm:ss")

630 sql = "select * from patientdetails where " & _
          "labnumber = '" & tLabNum & "'"
640 Set ds = New Recordset
650 RecOpenClientBB 0, ds, sql
660 If ds.EOF Then
670     ds.AddNew
680 Else
690     retval = False
700     If Trim$(UCase$(ds!Name & "")) <> Trim$(UCase$(txtName)) Then
710         retval = FlagMessage("Name", ds!Name & "", txtName)
720         If TimedOut Then Unload Me: Exit Function
730         If retval Then Exit Function
740     ElseIf Format(ds!DoB, "dd/mm/yyyy") <> Format(tDoB, "dd/mm/yyyy") Then
750         retval = FlagMessage("DoB", ds!DoB & "", tDoB)
760         If TimedOut Then Unload Me: Exit Function
770         If retval Then Exit Function
780     ElseIf Trim$(UCase$(ds!Patnum & "")) <> Trim$(UCase$(txtChart)) Then
790         retval = FlagMessage("Chart", ds!Patnum & "", txtChart)
800         If TimedOut Then Unload Me: Exit Function
810         If retval Then Exit Function
820     ElseIf Trim$(UCase$(ds!fGroup & "")) <> Trim$(UCase$(lstfg)) Then
            If Trim$(UCase$(ds!fGroup & "")) <> "" Then
830             retval = FlagMessage("Group", ds!fGroup & "", lstfg)
840             If TimedOut Then Unload Me: Exit Function
850             If retval Then Exit Function
            End If
860     End If
870 End If

880 ds!Kell = cmbKell

890 ds!Hold = Hold

900 ds!Checker = lblgrpchecker

910 If IsDate(txtSampleDate) And IsDate(txtSampleTime) Then
920     ds!SampleDate = Format(txtSampleDate & " " & txtSampleTime, "dd/MMM/yyyy hh:mm:ss")
930 ElseIf IsDate(txtSampleDate) Then
940     ds!SampleDate = Format(txtSampleDate, "dd/MMM/yyyy")
950 Else
960     ds!SampleDate = Null
970 End If
980 DateTimeRxd = Format(dtDateRxd, "dd/mmm/yyyy")
990 If imgUseTime.Picture = imgSquareTick.Picture Then
1000    DateTimeRxd = DateTimeRxd & " " & dtTimeRxd
1010 End If
1020 ds!DateReceived = Format(DateTimeRxd, "dd/mmm/yyyy HH:nn:ss")

1030 ds!LabNumber = tLabNum
1040 ds!Patnum = txtChart
1050 ds!AandE = tAandE
1060 ds!Typenex = tTypenex

    '980   If InStr(lblChartNumber, "Cavan") Then
1070 HospitalName = "C"
    '1000  ElseIf InStr(lblChartNumber, "Monaghan") Then
    '1010    HospitalName = "M"
    '1020  End If
1080 ds!Hospital = HospitalName
1090 ds!Name = txtName
1100 ds!PatSurName = txtSurname
1110 ds!PatForeName = txtForname
1120 ds!maiden = tMaiden
1130 ds!coombs = " "
1140 If IsDate(tDoB) Then
1150    ds!DoB = Format(tDoB, "dd/mmm/yyyy")
1160 Else
1170    ds!DoB = Null
1180 End If
1190 ds!Age = tAge
1200 ds!Ward = Left$(cWard, 20)
1210 ds!GP = Left$(cGP, 20)
1220 ds!Clinician = Left$(cClinician, 50)
1230 ds!Procedure = Left$(cProcedure, 50)
1240 ds!Conditions = Left$(cConditions, 50)
1250 ds!specialprod = Left$(cSpecial, 50)
1260 ds!Addr1 = tAddr(0)
1270 ds!Addr2 = tAddr(1)
1280 ds!Addr3 = tAddr(2)
1290 ds!addr4 = tAddr(3)
1300 image2grh pGroup, prh
1310 ds!PrevGroup = pGroup
1320 ds!previousrh = prh
1330 ds!Sex = Left$(lSex, 1)
1340 ds!Comment = tComment
1350 ds!fgsuggest = lblsuggestfg
1360 ds!edd = IIf(IsDate(tedd), tedd, Null)
1370 ds!SampleComment = Trim$(tSampleComment)

1380 s = ""

1390 If s <> "" And Len(s) <> 8 Then
1400    s = ""
1410    iMsg "Grouping Pattern incomplete!" & vbCrLf & "Not saved!", vbExclamation
1420    If TimedOut Then Unload Me: Exit Function
1430 End If


1440 ds!fgpattern = s

1450 ds!fGroup = lstfg

1460 ds!AIDR = tident
1470 ds!DateTime = TimeNow

1480 For n = 0 To 11
1490    If gDAT.Rows > (n \ 2) + 1 Then
1500        gDAT.col = (n Mod 2) + 1
1510        gDAT.row = (n \ 2) + 1
1520        ds("DAT" & Format$(n)) = gDAT.CellPicture = imgSquareCross.Picture
1530    End If
    'ds("DAT" & Format(n)) = optDAT(n)
1540 Next

1550 ds!Operator = UserCode
1560 ds!ampm = 0
1570 ds!Autolog = 0
1580 ds!IsPublic = IIf(optPubPri(0), 1, 0)
1590 ds.Update

1600 cmdSave.Enabled = False: bHold.Enabled = False

1610 SaveAutomaticEligibility

1620 ShowEligibility

1630 SaveDetails = True

1640 Exit Function

SaveDetails_Error:

    Dim strES As String
    Dim intEL As Integer

1650 intEL = Erl
1660 strES = Err.Description
1670 LogError "frmxmatch", "SaveDetails", intEL, strES, sql

End Function

Private Sub SuggestLabNum()

10  tLabNum = GetSetting("NetAcquire", "Transfusion6", "LastUsed", "1")

End Sub

Private Sub tAandE_LostFocus()

10  On Error GoTo tAandE_LostFocus_Error

20  lblPrevAdverse.Visible = False

30  If tAandE <> "" Then
40      LoadPatFromNCA "AANDE", tAandE
50  End If

60  If tAandE <> "" And txtName = "" Then
70      FillFromNumber "AandE", tAandE
80  End If

90  If CheckBadReaction(txtChart) Then
100     lblPrevAdverse.Visible = True
110 End If

120 lKnownAntibody.Caption = CheckPreviousABScreen(txtChart)
130 If lKnownAntibody.Caption <> "" Then
140     lKnownAntibody.Visible = True
150 End If

160 Exit Sub

tAandE_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

170 intEL = Erl
180 strES = Err.Description
190 LogError "frmxmatch", "tAandE_LostFocus", intEL, strES

End Sub


Private Sub FillFromNumber(ByVal Identifier As String, _
                           ByVal Number As String)


    Dim sn As Recordset
    Dim sql As String

    'lblPreviousGroup = ""

10  On Error GoTo FillFromNumber_Error

20  If Trim$(Number) = "" Then Exit Sub

30  sql = "select * from patientdetails where "
40  Select Case Identifier
        Case "MRN"
50          sql = sql & "patnum = '" & Number & "' "
60      Case "NOPAS"
70          sql = sql & "nopas = '" & Number & "' "
80      Case "AandE"
90          sql = sql & "AandE = '" & Number & "' "
100 End Select
110 sql = sql & "order by datetime"
120 Set sn = New Recordset
130 RecOpenServerBB 0, sn, sql
140 If sn.EOF Then Exit Sub
150 sn.MoveLast

160 txtChart = sn!Patnum & ""
170 If Trim$(txtChart) <> "" Then
180     cmdLab.Visible = True
190 Else
200     cmdLab.Visible = False
210 End If
220 txtName = sn!Name & ""
230 If Not IsNull(sn!DoB) Then
240     tDoB = Format(sn!DoB, "dd/mm/yyyy")
250 Else
260     tDoB = ""
270 End If
280 tAge = sn!Age & ""
290 cWard = sn!Ward & ""
300 cClinician = sn!Clinician & ""
310 cConditions = sn!Conditions & ""
320 cProcedure = sn!Procedure & ""
330 cSpecial = sn!specialprod & ""

340 tAddr(0) = sn("addr1") & ""
350 tAddr(1) = sn("addr2") & ""
    'lblPreviousGroup = GetPreviousGroup()
360 lSex = sn!Sex & ""
370 tComment = StripComment(sn!Comment & "")
380 If Trim$(sn!Age & "") <> "" Then
390     tAge = sn!Age
400 ElseIf tDoB <> "" Then
410     tAge = CalcAge(tDoB)
420 End If
430 If Not IsDate(sn!SampleDate & "") Then
440     txtSampleDate = ""
450     txtSampleTime = ""
460 Else
470     txtSampleDate = Format(sn!SampleDate, "dd/MM/yyyy")
480     If Format(sn!SampleDate, "HH:mm") <> "00:00" Then
490         txtSampleTime = Format(sn!SampleDate, "HH:nn:ss")
500     Else
510         txtSampleTime = ""
520     End If
530 End If
540 If IsNull(sn!DateReceived) Then
550     dtDateRxd = Format(Now, "dd/MM/yyyy")
560     dtTimeRxd = Format(Now, "HH:nn:ss")
570     Set imgUseTime.Picture = imgSquareCross.Picture
580     dtTimeRxd.Enabled = False
590 Else
600     dtDateRxd = Format(sn!DateReceived, "dd/MM/yyyy")
610     dtTimeRxd = Format(TimeValue(sn!DateReceived), "HH:nn:ss")
620     If Format(TimeValue(sn!DateReceived), "HH:mm") <> "00:00" Then
630         Set imgUseTime.Picture = imgSquareTick.Picture
640         dtTimeRxd.Enabled = True
650     Else
660         Set imgUseTime.Picture = imgSquareCross.Picture
670         dtTimeRxd.Enabled = True
680     End If
690 End If

700 cmdSave.Enabled = False: bHold.Enabled = False

710 Exit Sub

FillFromNumber_Error:

    Dim strES As String
    Dim intEL As Integer

720 intEL = Erl
730 strES = Err.Description
740 LogError "frmxmatch", "FillFromNumber", intEL, strES, sql

End Sub




Public Sub LoadPatFromNCA(ByVal strType As String, ByVal Value As String)

'strType is either "NOPAS" or "CHART" or "AANDE"

    Dim tbPatIF As Recordset
    Dim tbDemog As Recordset
    Dim sql As String
    Dim strPatientEntity As String
    Dim X As Long
    Dim NoDateTime As Boolean

10  On Error GoTo LoadPatFromNCA_Error

20  strPatientEntity = ""

30  sql = "Select * from PatientIFs where " & _
          strType & " = '" & Value & "' "
40  Set tbPatIF = New Recordset
50  RecOpenServer 0, tbPatIF, sql

60  sql = "select TOP 1 * from demographics where " & _
          strType & "= '" & Value & "' " & _
          "order by DateTimeDemographics desc"
70  Set tbDemog = New Recordset
80  RecOpenServer 0, tbDemog, sql

90  If tbPatIF.EOF And tbDemog.EOF Then
100     txtName = ""
110     tAddr(0) = ""
120     tAddr(1) = ""
130     lSex = ""
140     tDoB = ""
150     tAge = ""
160     cWard = "GP"
170     cClinician = ""
180     tComment = ""
190 ElseIf tbDemog.EOF Then
200     With tbPatIF
210         txtChart = !Chart & ""
220         tAandE = !AandE & ""
230         txtName = Initial2Upper(!PatName & "")
240         lSex = !Sex
250         tDoB = !DoB
260         tAge = CalcAge(tDoB)
270         cWard = !Ward & ""
280         cClinician = !Clinician & ""
290         tAddr(0) = Initial2Upper(!Address0 & "")
300         tAddr(1) = Initial2Upper(!Address1 & "")
310     End With
320 ElseIf tbPatIF.EOF Then
330     txtName = tbDemog!PatName & ""
340     tAddr(0) = tbDemog!Addr0 & ""
350     tAddr(1) = tbDemog!Addr1 & ""
360     Select Case tbDemog!Sex & ""
            Case "M": lSex = "Male"
370         Case "F": lSex = "Female"
380     End Select
390     txtChart = tbDemog!Chart & ""
400     tAandE = tbDemog!AandE & ""
410     tAge = tbDemog!Age & ""
420     tDoB = Format$(tbDemog!DoB, "dd/mm/yyyy")
430     If IsDate(tDoB) Then
440         tAge = CalcAge(tDoB)
450     End If
460     cWard = tbDemog!Ward & ""
470     cClinician = tbDemog!Clinician & ""
480 Else
490     NoDateTime = False
500     If IsNull(tbDemog!DateTimeDemographics) Or IsNull(tbPatIF!DateTimeAmended) Then
510         NoDateTime = True
520     End If
530     X = -1
540     If Not NoDateTime Then
550         X = DateDiff("h", tbDemog!DateTimeDemographics, tbPatIF!DateTimeAmended)
560     End If
570     If X < 0 Or IsNull(X) Then
580         txtName = tbDemog!PatName & ""
590         tAddr(0) = tbDemog!Addr0 & ""
600         tAddr(1) = tbDemog!Addr1 & ""
610         Select Case tbDemog!Sex & ""
                Case "M": lSex = "Male"
620             Case "F": lSex = "Female"
630         End Select
640         txtChart = tbDemog!Chart & ""
650         tAandE = tbDemog!AandE & ""
660         tAge = tbDemog!Age & ""
670         tDoB = Format$(tbDemog!DoB, "dd/mm/yyyy")
680         If IsDate(tDoB) Then
690             tAge = CalcAge(tDoB)
700         End If
710         cWard = tbDemog!Ward & ""
720         cClinician = tbDemog!Clinician & ""
730     Else
740         With tbPatIF
750             txtChart = !Chart & ""
760             tAandE = !AandE & ""
770             txtName = Initial2Upper(!PatName & "")
780             lSex = !Sex
790             tDoB = !DoB
800             tAge = CalcAge(tDoB)
810             cWard = !Ward & ""
820             cClinician = !Clinician & ""
830             tAddr(0) = Initial2Upper(!Address0 & "")
840             tAddr(1) = Initial2Upper(!Address1 & "")
850         End With
860     End If
870 End If

880 Exit Sub

LoadPatFromNCA_Error:

    Dim strES As String
    Dim intEL As Integer

890 intEL = Erl
900 strES = Err.Description
910 LogError "frmxmatch", "LoadPatFromNCA", intEL, strES, sql

End Sub

Private Sub taddr_LostFocus(Index As Integer)

10  tAddr(Index) = Initial2Upper(tAddr(Index))
20  If Index = 3 Or Index = 4 Then tComment.SetFocus

End Sub

Private Sub tComment_Click()

    Dim f As Form

10  Dept = XMATCH

20  Set f = New frmRemarks

30  f.Comment = tComment
40  f.Heading = "Remarks Entry"
50  f.Show 1
60  tComment = f.Comment
70  Unload f

80  Set f = Nothing

90  labnumberfound = tLabNum

End Sub

Private Sub tcomment_LostFocus()

10  If UCase$(tComment) = "N" Then tComment = "No Irregular Agglutinins Found"

End Sub

Private Sub tLabNum_GotFocus()

10  If Trim$(tLabNum) = "" Then Exit Sub
20  If blnShowingBox Then
30      blnShowingBox = False
40      Exit Sub
50  End If

60  If cmdSave.Enabled Then
70      blnShowingBox = True
80      Answer = iMsg("Save Details?", vbYesNo + vbQuestion)
90      If TimedOut Then Unload Me: Exit Sub
100     If Answer = vbYes Then
110         cmdSave.Value = True
120         cmdSave.Enabled = False: bHold.Enabled = False
130     End If
140 End If

End Sub

Public Sub tlabnum_LostFocus()

10  If blnShowingBox Then Exit Sub
20  lKnownAntibody.Visible = False  '(Nk)
    'Zyam 2-12-24
30  LoadLabNumber
    'Zyam 2-12-24

40  If Trim$(txtChart) <> "" Then
50      cmdLab.Visible = True
60  Else
70      cmdLab.Visible = False
80  End If
    FillFromNumber "MRN", tLabNum.Text

End Sub

Private Sub tmaiden_LostFocus()

10  tMaiden = Initial2Upper(tMaiden)

End Sub

Private Sub tSampleComment_KeyPress(KeyAscii As Integer)

10  cmdSave.Enabled = True: bHold.Enabled = True

End Sub


Private Sub txtNOPAS_LostFocus()

10  On Error GoTo txtNOPAS_LostFocus_Error

20  lblPrevAdverse.Visible = False

30  txtNOPAS = Trim$(txtNOPAS)

40  If txtNOPAS <> "" Then
50      LoadPatFromNCA "NOPAS", txtNOPAS
60  End If

70  If txtNOPAS <> "" And txtName = "" Then
80      FillFromNumber "NOPAS", txtNOPAS
90  End If

100 If CheckBadReaction(txtChart) Then
110     lblPrevAdverse.Visible = True
120 End If

130 lKnownAntibody.Caption = CheckPreviousABScreen(txtChart)
140 If lKnownAntibody.Caption <> "" Then
150     lKnownAntibody.Visible = True
160 End If

170 Exit Sub

txtNOPAS_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmxmatch", "txtNOPAS_LostFocus", intEL, strES

End Sub


Private Sub txtSampleDate_DblClick()

10  txtSampleDate = Format(Now, "dd/mm/yyyy")
20  txtSampleTime = ""

30  cmdSave.Enabled = True: bHold.Enabled = True

End Sub

Private Sub txtSampleDate_LostFocus()

10  txtSampleDate = Convert62Date(txtSampleDate, BACKWARD)

End Sub

Private Sub tDoB_LostFocus()

10  tDoB = Convert62Date(tDoB, BACKWARD)
20  tAge = CalcAge(tDoB)

End Sub

Private Sub txtChart_Change()

10  SearchCriteria = txtChart
20  If Not tLabNum.Locked Then
30      cmdSearch.Enabled = True
40  End If

End Sub

Private Sub txtChart_LostFocus()

    Dim sn As Recordset
    Dim sql As String
    Dim strPatientEntity As String
    Dim tbIF As Recordset

10  On Error GoTo txtChart_LostFocus_Error

20  txtChart = Trim$(txtChart)
30  bhistory.BackColor = &H8000000F
40  lKnownAntibody.Visible = False

50  If Trim$(txtChart) = "" Then Exit Sub

60  lKnownAntibody = CheckPreviousABScreen(txtChart)
70  If lKnownAntibody <> "" Then
80      lKnownAntibody.Visible = True
90  End If

100 lblGroupDisagree.Visible = False
110 If Not GroupHistoryOK() Then
120     lblGroupDisagree.Visible = True
130 End If

140 If Trim$(txtName) <> "" Then Exit Sub

150 strPatientEntity = GetOptionSetting("IMSPASPATIENTENTITY", "CGH")

160 If GetOptionSetting("IMSPASOPERATIONAL", 1) <> "1" Then    'Using PAS

170     sql = "SELECT DISTINCT P.Chart, P.PatName, P.PatSurName, P.PatForeName, " & _
              "CASE P.Sex WHEN 'M' THEN 'Male' WHEN 'F' THEN 'Female' ELSE '' END Sex, " & _
              "P.DoB, W.Text Ward, " & _
              "C.Text Clinician, Address0, Address1, Entity, Episode, " & _
              "RegionalNumber, DateTimeAmended, NewEntry, AandE, " & _
              "MRN, AdmitDate ,G.Text GP " & _
              "FROM PatientIFs P " & _
              "LEFT JOIN Wards W ON W.Code = P.Ward " & _
              "LEFT JOIN Clinicians C ON C.Code = P.Clinician " & _
              "LEFT JOIN GPs G ON G.Code = P.GP " & _
              "WHERE Chart = '" & txtChart & "' " & _
              "AND Entity = '" & strPatientEntity & "'"

180     Set tbIF = New Recordset
190     RecOpenClient 0, tbIF, sql
200     If Not tbIF.EOF Then
210         txtName = Initial2Upper(tbIF!PatName & "")
220         txtSurname = Trim$(tbIF!PatSurName & "")
230         txtForname = Trim$(tbIF!PatForeName & "")
240         lSex = tbIF!Sex & ""
250         tDoB = tbIF!DoB
260         cWard = tbIF!Ward & ""
270         cClinician = tbIF!Clinician & ""
280         cGP = tbIF!GP & ""
290         tAddr(0) = Initial2Upper(tbIF!Address0 & "")
300         tAddr(1) = Initial2Upper(tbIF!Address1 & "")
310         bhistory.BackColor = &H8080FF
320     Else
330         sql = "SELECT TOP 1 Name, PatForeName, PatSurName, Sex, DoB, Ward, GP, Clinician, Addr1, Addr2, Addr3, Addr4, " & _
                  "CASE LEFT(Sex, 1) WHEN 'M' THEN 'Male' WHEN 'F' THEN 'Female' ELSE '' END Sex, " & _
                  "Conditions, Age, [Procedure], SpecialProd, Maiden, PrevGroup, PreviousRh, Comment, EDD " & _
                  "FROM PatientDetails WHERE " & _
                  "PatNum = '" & txtChart & "' " & _
                  "AND Hospital = 'C' " & _
                  "ORDER BY DateTime desc"
340         Set sn = New Recordset
350         RecOpenServerBB 0, sn, sql
360         If sn.EOF Then Exit Sub

370         txtName = sn!Name & ""
380         txtSurname = Trim$(sn!PatSurName & "")
390         txtForname = Trim$(sn!PatForeName & "")
400         lSex = sn!Sex & ""
410         If Not IsNull(sn!DoB) Then
420             tDoB = Format(sn!DoB, "dd/mm/yyyy")
430         Else
440             tDoB = ""
450         End If
460         cWard = sn!Ward & ""
470         cGP = sn!GP & ""
480         cClinician = sn!Clinician & ""
490         tAddr(0) = sn!Addr1 & ""
500         tAddr(1) = sn!Addr2 & ""

510         cConditions = sn!Conditions & ""
520         tAge = sn!Age & ""
530         cProcedure = sn!Procedure & ""
540         cSpecial = sn!specialprod & ""
550         tMaiden = sn!maiden & ""
560         tAddr(2) = sn!Addr3 & ""
570         tAddr(3) = sn!addr4 & ""
580         grh2image sn!PrevGroup & "", sn!previousrh & ""
590         tComment = StripComment(sn!Comment & "")
600         If Not IsNull(sn!edd) Then
610             tedd = Format(sn!edd, "dd/mm/yyyy")
620         Else
630             tedd = ""
640         End If

650     End If
660 End If

670 Exit Sub

txtChart_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

680 intEL = Erl
690 strES = Err.Description
700 LogError "frmxmatch", "txtChart_LostFocus", intEL, strES, sql

End Sub

Private Sub txtName_Change()

10  SearchCriteria = txtName
20  If Not tLabNum.Locked Then
30      cmdSearch.Enabled = True
40  End If

End Sub

Private Sub txtName_LostFocus()

    Dim strName As String
    Dim strSex As String

10  If lSex <> "" Then Exit Sub

20  strName = txtName
30  strSex = lSex

40  NameLostFocus strName, strSex

50  txtName = strName
60  lSex = strSex

End Sub

Public Property Let LoadingFlag(ByVal vNewValue As Variant)

10  Loading = vNewValue

End Property

Private Sub udLabNum_GotFocus()

10  If Trim$(tLabNum) = "" Then Exit Sub

20  If cmdSave.Enabled Then
30      Answer = iMsg("Save Details?", vbYesNo + vbQuestion)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbYes Then
60          cmdSave.Value = True
70          cmdSave.Enabled = False: bHold.Enabled = False
80      End If
90  End If

End Sub


Private Sub udLabNum_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10  LoadLabNumber

End Sub


Private Function LockDemographics(LockIt As Boolean) As Boolean
10  If LockIt Then
20      cmdLockDemo.Caption = "Unlock"
30      cmdLockDemo.BackColor = vbRed
40  Else
50      cmdLockDemo.Caption = "Lock"
60      cmdLockDemo.BackColor = vbGreen
70  End If
80  lSex.Enabled = Not LockIt
90  txtChart.Locked = LockIt
100 tAandE.Locked = LockIt
110 tDoB.Locked = LockIt
120 tTypenex.Locked = LockIt
130 txtName.Locked = LockIt
140 tAddr(0).Locked = LockIt
150 tAddr(1).Locked = LockIt
160 tAddr(2).Locked = LockIt
170 tAddr(3).Locked = LockIt
180 cWard.Locked = LockIt
190 cClinician.Locked = LockIt
200 FramePP.Enabled = Not LockIt
End Function

Private Function ValidForConfirmation() As Boolean

10  On Error GoTo ValidForConfirmation_Error

20  If tLabNum = "" Then
30      iMsg "Please enter sample number first", vbInformation
40      If TimedOut Then Unload Me: Exit Function
50      tLabNum.SetFocus
60      ValidForConfirmation = False
70      Exit Function
80  End If

90  If txtChart = "" And tAandE = "" Then
100     iMsg "Please enter MRN or A/E number first", vbInformation
110     If TimedOut Then Unload Me: Exit Function
120     txtChart.SetFocus
130     ValidForConfirmation = False
140     Exit Function
150 End If

160 If Len(Trim$(tTypenex)) <> 8 Then
170     iMsg "Please enter typenex first", vbInformation
180     If TimedOut Then Unload Me: Exit Function
190     tTypenex.SetFocus
200     ValidForConfirmation = False
210     Exit Function
220 End If

230 If txtName = "" Then
240     iMsg "Please enter patient name first", vbInformation
250     If TimedOut Then Unload Me: Exit Function
260     txtName.SetFocus
270     ValidForConfirmation = False
280     Exit Function
290 End If

300 If lSex.Caption = "" Then
310     iMsg "Please select patient sex first", vbInformation
320     If TimedOut Then Unload Me: Exit Function
330     ValidForConfirmation = False
340     Exit Function
350 End If

360 ValidForConfirmation = True
370 Exit Function

ValidForConfirmation_Error:

    Dim strES As String
    Dim intEL As Integer

380 intEL = Erl
390 strES = Err.Description
400 LogError "frmxmatch", "ValidForConfirmation", intEL, strES

End Function

Private Sub LockUnlockConfirmedMallow()

    Dim EnableFlag As Boolean

10  On Error GoTo LockUnlockConfirmedMallow_Error

20  If PatientDetailsConfirmed(tLabNum) Then
30      cmdValidate.Caption = "UnConfirm"
40      EnableFlag = False

50  Else
60      cmdValidate.Caption = "Confirm Details"
70      EnableFlag = True

80  End If

90  bPrintForm.Enabled = Not EnableFlag
100 btnprint.Enabled = Not EnableFlag
110 cmdLockDemo.Enabled = EnableFlag
120 txtChart.Enabled = EnableFlag
130 tAandE.Enabled = EnableFlag
140 txtNOPAS.Enabled = EnableFlag
150 tTypenex.Enabled = EnableFlag
160 txtName.Enabled = EnableFlag
170 tMaiden.Enabled = EnableFlag
180 tAddr(0).Enabled = EnableFlag
190 tAddr(1).Enabled = EnableFlag
200 tAddr(2).Enabled = EnableFlag
210 tAddr(3).Enabled = EnableFlag
220 tDoB.Enabled = EnableFlag
230 tAge.Enabled = EnableFlag
240 lSex.Enabled = EnableFlag
250 tComment.Enabled = EnableFlag
260 cWard.Enabled = EnableFlag
270 cClinician.Enabled = EnableFlag
280 cGP.Enabled = EnableFlag
290 cConditions.Enabled = EnableFlag
300 cProcedure.Enabled = EnableFlag
310 cSpecial.Enabled = EnableFlag
320 tSampleComment.Enabled = EnableFlag
330 tedd.Enabled = EnableFlag
340 txtSampleDate.Enabled = EnableFlag
350 txtSampleTime.Enabled = EnableFlag
360 dtDateRxd.Enabled = EnableFlag
370 dtTimeRxd.Enabled = True
380 imgUseTime.Enabled = EnableFlag

390 Exit Sub

LockUnlockConfirmedMallow_Error:

    Dim strES As String
    Dim intEL As Integer

400 intEL = Erl
410 strES = Err.Description
420 LogError "frmxmatch", "LockUnlockConfirmedMallow", intEL, strES

End Sub


Private Function isIDPanelPresent4SID(ByVal strLabNum As String) As Boolean
    Dim sql As String
    Dim sn As Recordset

10  On Error GoTo isIDPanelPresent4SID_Error

20  isIDPanelPresent4SID = False

30  sql = "SELECT Sampleid FROM AnalyserIDPanels WHERE Sampleid = '" & strLabNum & "'"

40  Set sn = New Recordset
50  RecOpenServerBB 0, sn, sql
60  If Not sn.EOF Then
70      isIDPanelPresent4SID = True
80  End If

90  Exit Function

isIDPanelPresent4SID_Error:

    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmxmatch", "isIDPanelPresent4SID", intEL, strES, sql

End Function







