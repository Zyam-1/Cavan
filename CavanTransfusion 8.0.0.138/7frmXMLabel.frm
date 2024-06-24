VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "ComCt232.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmXMLabel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cross-Match Labels"
   ClientHeight    =   8925
   ClientLeft      =   525
   ClientTop       =   660
   ClientWidth     =   14430
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "7frmXMLabel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8925
   ScaleWidth      =   14430
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2985
      Left            =   8040
      ScaleHeight     =   2925
      ScaleWidth      =   2970
      TabIndex        =   52
      Top             =   465
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.CommandButton cmdNotPrinted 
      Caption         =   "Mark as Not Printed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   11700
      Picture         =   "7frmXMLabel.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7620
      Width           =   1275
   End
   Begin VB.CommandButton cmdReprintPDF 
      Caption         =   "Reprint P&DF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   11700
      MaskColor       =   &H8000000F&
      Picture         =   "7frmXMLabel.frx":240C
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6150
      Width           =   1275
   End
   Begin VB.Timer tmrCourier 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   12705
      Top             =   4830
   End
   Begin VB.Frame fraCourier 
      Caption         =   "Communicating with Courier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11160
      TabIndex        =   46
      Top             =   4710
      Visible         =   0   'False
      Width           =   2355
      Begin MSComctlLib.ProgressBar pbCourier 
         Height          =   225
         Left            =   60
         TabIndex        =   47
         Top             =   240
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
         Max             =   300
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton cmdPDF 
      Caption         =   "Print P&DF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   11700
      MaskColor       =   &H8000000F&
      Picture         =   "7frmXMLabel.frx":284E
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3870
      Width           =   1275
   End
   Begin VB.CommandButton bSelect 
      Caption         =   "Select &All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   2040
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "Serum From Specimen Labelled"
      Height          =   2865
      Left            =   1410
      TabIndex        =   10
      Top             =   60
      Width           =   6615
      Begin VB.TextBox txtForname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6300
         MaxLength       =   30
         TabIndex        =   54
         Tag             =   "Name"
         Top             =   510
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtSurname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6300
         MaxLength       =   30
         TabIndex        =   53
         Tag             =   "Name"
         Top             =   300
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtAddr2 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2070
         Width           =   4875
      End
      Begin VB.TextBox tSex 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   5460
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1470
         Width           =   735
      End
      Begin VB.TextBox tClinician 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   870
         Width           =   4875
      End
      Begin VB.TextBox tAandE 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1470
         Width           =   1035
      End
      Begin VB.TextBox tSampleDate 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4500
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2370
         Width           =   1695
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   270
         Width           =   4875
      End
      Begin VB.TextBox tward 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   16
         Top             =   570
         Width           =   4875
      End
      Begin VB.TextBox tdob 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4260
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1470
         Width           =   1155
      End
      Begin VB.TextBox tgroup 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Top             =   2370
         Width           =   1395
      End
      Begin VB.TextBox txtChart 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1470
         Width           =   1155
      End
      Begin VB.TextBox txtAddr1 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   12
         Top             =   1770
         Width           =   4875
      End
      Begin VB.TextBox tcomment 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1170
         Width           =   4875
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Clinician"
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
         Left            =   690
         TabIndex        =   38
         Top             =   900
         Width           =   585
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "A/E"
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
         Left            =   2520
         TabIndex        =   34
         Top             =   1530
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sample Date"
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
         Left            =   3540
         TabIndex        =   25
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "MRN (Chart No.)"
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
         Left            =   90
         TabIndex        =   24
         Top             =   1500
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         Left            =   840
         TabIndex        =   23
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
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
         Left            =   870
         TabIndex        =   22
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DoB"
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
         Left            =   3930
         TabIndex        =   21
         Top             =   1500
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Address"
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
         Left            =   720
         TabIndex        =   20
         Top             =   1830
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Group/Rh"
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
         Left            =   570
         TabIndex        =   19
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Comment"
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
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   1200
         Width           =   660
      End
   End
   Begin VB.CommandButton bprint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   11700
      Picture         =   "7frmXMLabel.frx":2C90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2430
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   11475
      TabIndex        =   3
      Top             =   1140
      Width           =   1815
      Begin ComCtl2.UpDown udCopiesPlus 
         Height          =   285
         Left            =   1050
         TabIndex        =   9
         Top             =   660
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "tCopiesPlus"
         BuddyDispid     =   196643
         OrigLeft        =   1500
         OrigTop         =   960
         OrigRight       =   1905
         OrigBottom      =   1200
         Max             =   9
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox tCopiesPlus 
         Height          =   285
         Left            =   1050
         TabIndex        =   8
         Text            =   "1"
         Top             =   330
         Width           =   480
      End
      Begin VB.TextBox tCopies 
         Height          =   285
         Left            =   270
         MaxLength       =   1
         TabIndex        =   6
         Text            =   "1"
         Top             =   330
         Width           =   480
      End
      Begin ComCtl2.UpDown udCopies 
         Height          =   285
         Left            =   270
         TabIndex        =   5
         Top             =   660
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         _Version        =   327681
         Value           =   2
         BuddyControl    =   "tCopies"
         BuddyDispid     =   196644
         OrigLeft        =   1080
         OrigTop         =   1200
         OrigRight       =   1425
         OrigBottom      =   1440
         Max             =   9
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "+"
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   120
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   3900
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   3413
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
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"7frmXMLabel.frx":32FA
   End
   Begin VB.TextBox tcompat 
      Height          =   285
      Left            =   1860
      TabIndex        =   1
      Text            =   "Electronically issued as compatible"
      Top             =   3570
      Width           =   5115
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   13095
      Picture         =   "7frmXMLabel.frx":3393
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7620
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   1410
      TabIndex        =   32
      Top             =   2940
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid gPrinted 
      Height          =   2655
      Left            =   120
      TabIndex        =   48
      Top             =   6150
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   4683
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
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"7frmXMLabel.frx":425D
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Already Printed"
      Height          =   195
      Left            =   120
      TabIndex        =   49
      Top             =   5910
      Width           =   1305
   End
   Begin VB.Label lAB 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2460
      TabIndex        =   44
      Top             =   3240
      Width           =   4515
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Antibodies"
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
      Index           =   0
      Left            =   1710
      TabIndex        =   43
      Top             =   3270
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Operator"
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
      Index           =   1
      Left            =   11265
      TabIndex        =   42
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lOp 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11925
      TabIndex        =   41
      Top             =   210
      Width           =   1350
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Typenex"
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
      Index           =   1
      Left            =   11235
      TabIndex        =   36
      Top             =   780
      Width           =   615
   End
   Begin VB.Label lTypenex 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   11925
      TabIndex        =   35
      Top             =   750
      Width           =   1350
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "RED"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   30
      Top             =   2970
      Width           =   825
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Units in             will Print"
      Height          =   645
      Left            =   180
      TabIndex        =   29
      Top             =   2730
      Width           =   915
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   420
      Picture         =   "7frmXMLabel.frx":42F3
      Top             =   3390
      Width           =   480
   End
   Begin VB.Label lLabNumber 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11925
      TabIndex        =   28
      Top             =   480
      Width           =   1350
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Lab Number"
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
      Index           =   0
      Left            =   10995
      TabIndex        =   27
      Top             =   510
      Width           =   870
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   6990
      Picture         =   "7frmXMLabel.frx":4735
      Stretch         =   -1  'True
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1380
      Picture         =   "7frmXMLabel.frx":4B77
      Stretch         =   -1  'True
      Top             =   3570
      Width           =   480
   End
End
Attribute VB_Name = "frmXMLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String

Private Activated As Boolean
Private Sub ExecPrintLabel()

          Dim n As Integer
          Dim Y As Integer

10    On Error GoTo ExecPrintLabel_Error

20    ReDim RowNumbersToPrint(1 To 1) As Integer

30    n = 0
40    g.col = 0
50    For Y = 1 To g.Rows - 1
60      g.row = Y
70      If g.CellBackColor = vbRed Then
80          n = n + 1
90          ReDim Preserve RowNumbersToPrint(1 To n) As Integer
100         RowNumbersToPrint(n) = Y
110     End If
120   Next

130   If n = 0 Then
140     iMsg "Nothing to do!" & vbCr & _
             "Click on units to print.", vbExclamation
150     If TimedOut Then Unload Me: Exit Sub
160     Exit Sub
170   End If

180   If Val(tCopies) < 1 Then
190     iMsg "Select number of copies required.", vbInformation
200     If TimedOut Then Unload Me: Exit Sub
210     Exit Sub
220   End If

230   If (Val(tCopies) + Val(tCopiesPlus)) > 8 Then
240     iMsg "Maximum 8 Labels allowed.", vbCritical
250     If TimedOut Then Unload Me: Exit Sub
260     Exit Sub
270   End If

280   bprint.Caption = "Printing..."
290   PrintLabel RowNumbersToPrint
300   bprint.Caption = "Print"

310   Exit Sub

ExecPrintLabel_Error:

          Dim strES As String
          Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmXMLabel", "ExecPrintLabel", intEL, strES

End Sub


Private Sub ShowCourierMessage()

          Dim C As Control

10    On Error GoTo ShowCourierMessage_Error

20    For Each C In Me.Controls
30      C.Enabled = False
40    Next

50    fraCourier.Visible = True
60    fraCourier.Enabled = True
70    pbCourier.Value = 0
80    tmrCourier.Enabled = True

90    Exit Sub

ShowCourierMessage_Error:

          Dim strES As String
          Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmXMLabel", "ShowCourierMessage", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bprint_Click()

10    ExecPrintLabel

End Sub

Private Sub bSelect_Click()

          Dim Y As Integer

10    g.col = 0

20    If bSelect.Caption = "Select All" Then
30      bSelect.Caption = "De-Select All"
40      For Y = 1 To g.Rows - 1
50          g.row = Y
60          g.CellBackColor = vbRed
70          g.CellForeColor = vbYellow
80      Next
90    Else
100     bSelect.Caption = "Select All"
110     For Y = 1 To g.Rows - 1
120         g.row = Y
130         g.CellBackColor = &H80000018
140         g.CellForeColor = &H8000000D
150     Next
160   End If

End Sub


Private Sub cmdNotPrinted_Click()

          Dim done As Boolean
          Dim Y As Integer

10    On Error GoTo cmdNotPrinted_Click_Error

20    done = False

30    gPrinted.col = 0
40    For Y = 1 To gPrinted.Rows - 1
50      gPrinted.row = Y
60      If gPrinted.CellBackColor = vbRed Then
70          done = True
80          MarkAsNotPrinted lLabNumber, gPrinted.TextMatrix(Y, 0), gPrinted.TextMatrix(Y, 2), ProductBarCodeFor(gPrinted.TextMatrix(Y, 3))
90      End If
100   Next

110   If Not done Then
120     iMsg "Nothing to do!" & vbCr & _
             "Select units.", vbExclamation
130     If TimedOut Then Unload Me: Exit Sub
140   Else
150     FillDetails
160   End If

170   Exit Sub

cmdNotPrinted_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmXMLabel", "cmdNotPrinted_Click", intEL, strES

End Sub

Private Sub cmdPDF_Click()

      Dim n As Integer
      Dim Y As Integer
      Dim tb As Recordset
      Dim sql As String
      Dim DeReservationDateTime As String
      Dim HoldFor As String
      Dim HoldForLG_OCTAPLAS As String
      Dim Generic As String
      Dim i As Integer
      Dim MSG As udtRS
      Dim BPs As New BatchProducts
      Dim BP As BatchProduct

10    On Error GoTo cmdPDF_Click_Error

20    ReDim RowNumbersToPrint(1 To 1) As Integer


30    If iMsg(vbCrLf & "Print label with comment:" & vbCrLf & vbCrLf & "< " & tcompat & " >" & "?", vbYesNo) = vbNo Then
40        Exit Sub
50    End If


60    HoldFor = GetOptionSetting("TransfusionHoldFor", "72")
70    HoldForLG_OCTAPLAS = GetOptionSetting("TransfusionHoldForLG_OCTAPLAS", "24")

80    n = 0
90    g.col = 0
100   For Y = 1 To g.Rows - 1
110       g.row = Y
120       If g.CellBackColor = vbRed Then
130           n = n + 1
140           ReDim Preserve RowNumbersToPrint(1 To n) As Integer
150           RowNumbersToPrint(n) = Y
160       End If
170   Next

180   If n = 0 Then
190       iMsg "Nothing to do!" & vbCr & _
               "Click on units to print.", vbExclamation
200       If TimedOut Then Unload Me: Exit Sub
210       Exit Sub
220   End If

230   cmdPDF.Caption = "Printing..."
240   For Y = 1 To UBound(RowNumbersToPrint)

250       PrintPDFCavan g.TextMatrix(RowNumbersToPrint(Y), 0), txtName, txtChart, tDoB, _
                        tward, tgroup, _
                        g.TextMatrix(RowNumbersToPrint(Y), 3), _
                        g.TextMatrix(RowNumbersToPrint(Y), 1), _
                        g.TextMatrix(RowNumbersToPrint(Y), 2), _
                        tcompat, _
                        tSampleDate, tSex, tgroup, txtSurname, txtForname

          '*******************************
          'Log unit received activity in courier interface requests table to update
          'status in blood courier management system by Blood Courier Interface
          'Implemented Site: CAVAN General Hospital
260       Generic = ProductGenericFor(ProductBarCodeFor(g.TextMatrix(RowNumbersToPrint(Y), 3)))
270       If UCase(Generic) = "RED CELLS" Or _
             UCase(Generic) = "PLATELETS" Or _
             UCase(Generic) = "LG OCTAPLAS" Then
              'Send RS signal
280           sql = "SELECT TOP 1 DateTime FROM Product WHERE " & _
                    "ISBT128 = '" & g.TextMatrix(RowNumbersToPrint(Y), 0) & "' " & _
                    "AND Event = 'X' OR Event = 'K' ORDER BY DateTime DESC"
290           Set tb = New Recordset
300           RecOpenClientBB 0, tb, sql
310           If Not tb.EOF Then
320               If UCase(Generic) = "LG OCTAPLAS" Then
330                 DeReservationDateTime = DateAdd("h", Val(HoldForLG_OCTAPLAS), Now)
340               Else
350                 DeReservationDateTime = DateAdd("h", Val(HoldFor), tSampleDate)
360               End If
370           Else
380               DeReservationDateTime = ""
390           End If

400           If IsDate(DeReservationDateTime) And IsDate(g.TextMatrix(RowNumbersToPrint(Y), 2)) Then
410               If DateDiff("D", Format$(DeReservationDateTime, "dd/MMM/yyyy"), Format$(g.TextMatrix(RowNumbersToPrint(Y), 2), "dd/MMM/yyyy")) < 0 Then
420                   DeReservationDateTime = Format$(g.TextMatrix(RowNumbersToPrint(Y), 2), "dd/MMM/yyyy hh:mm:ss")    '& " 23:59:59"
430               End If
440           End If

              'Cavan requested the following changes below
              'A "RTS" message should be send prior to the "RS" message
              '*******************************
              'Log unit movement activity in courier interface requests table to update
              'status in blood courier management system by Blood Courier Interface
              'Implemented Site: CAVAN General Hospital
              'Send RTS signal

450           With MSG
460               .UnitNumber = g.TextMatrix(RowNumbersToPrint(Y), 0)
470               .ProductCode = ProductBarCodeFor(g.TextMatrix(RowNumbersToPrint(Y), 3))
480               .UnitExpiryDate = g.TextMatrix(RowNumbersToPrint(Y), 2)
490               .StorageLocation = strBTCourier_StorageLocation_StockFridge
500               .ActionText = "Return to Stock"
510               .UserName = UserName
520           End With
530           LogCourierInterface "RTS", MSG

540           ShowCourierMessage

              '*******************************
              'Send "RS" message

550           With MSG
560               .UnitNumber = g.TextMatrix(RowNumbersToPrint(Y), 0)
570               .ProductCode = ProductBarCodeFor(g.TextMatrix(RowNumbersToPrint(Y), 3))

580               If Generic = "Platelets" Then
590                   .StorageLocation = strBTCourier_StorageLocation_RoomTempIssueFridge
600               Else
610                   .StorageLocation = strBTCourier_StorageLocation_HemoSafeFridge
620               End If
630               .UnitExpiryDate = g.TextMatrix(RowNumbersToPrint(Y), 2)
640               .UnitGroup = g.TextMatrix(RowNumbersToPrint(Y), 1)
650               .StockComment = ""
660               .Chart = txtChart
670               .PatientHealthServiceNumber = ""
680               If txtSurname & txtForname = "" Then
690                   i = InStr(1, txtName, " ")
700                   If i = 0 Then
710                       .ForeName = txtName
720                       .SurName = ""
730                   Else
740                       .SurName = Left$(txtName, i - 1)
750                       .ForeName = Mid$(txtName, i + 1, Len(txtName))
760                   End If
770               Else
780                   .SurName = txtSurname
790                   .ForeName = txtForname
800               End If
810               If tDoB <> "" Then
820                   .DoB = tDoB
830               End If
840               .Sex = Left$(tSex, 1)
850               .PatientGroup = tgroup
860               .DeReservationDateTime = Format(DeReservationDateTime, "dd-MMM-yyyy hh:mm:ss")
870               .ActionText = "Reserve Stock"
880               .UserName = UserName
890           End With
900           LogCourierInterface "RS3", MSG
910       End If

920       If Trim$(g.TextMatrix(RowNumbersToPrint(Y), 6)) <> "" Then    'its a batch product
930           BPs.LoadSpecificIdentifierLatest g.TextMatrix(RowNumbersToPrint(Y), 6)
940           If BPs.Count > 0 Then
950               Set BP = BPs.Item(1)
960               BP.LabelPrinted = 1
970               BP.Comment = "Label Printed"
980               BPs.Update BP
990           End If

              'BTC - Batch

              'Cavan requested the following changes below
              'A "RTS" message should be send prior to the "RS" message
              '*******************************
              'Log unit movement activity in courier interface requests table to update
              'status in blood courier management system by Blood Courier Interface
              'Implemented Site: CAVAN General Hospital
              'Send RTS signal

1000          With MSG
1010              .UnitNumber = g.TextMatrix(RowNumbersToPrint(Y), 6)
1020              .ProductCode = g.TextMatrix(RowNumbersToPrint(Y), 3)    'ProductBarCodeFor(g.TextMatrix(RowNumbersToPrint(y), 3))
1030              .UnitExpiryDate = g.TextMatrix(RowNumbersToPrint(Y), 2)
1040              .StorageLocation = strBTCourier_StorageLocation_StockFridge
1050              .ActionText = "Return to Stock"
1060              .UserName = UserName
1070          End With
1080          LogCourierInterface "RTS", MSG

1090          ShowCourierMessage

              '*******************************
              'Send "RS" message

1100          With MSG
1110              .UnitNumber = g.TextMatrix(RowNumbersToPrint(Y), 6)
1120              .ProductCode = g.TextMatrix(RowNumbersToPrint(Y), 3)    'ProductBarCodeFor(g.TextMatrix(RowNumbersToPrint(y), 3))

1130              If Generic = "" Then
1140                  .StorageLocation = strBTCourier_StorageLocation_RoomTempIssueFridge
1150              Else
1160                  .StorageLocation = strBTCourier_StorageLocation_HemoSafeFridge
1170              End If
1180              .UnitExpiryDate = g.TextMatrix(RowNumbersToPrint(Y), 2)
1190              .UnitGroup = g.TextMatrix(RowNumbersToPrint(Y), 1)
1200              .StockComment = ""
1210              .Chart = txtChart
1220              .PatientHealthServiceNumber = ""
1230              If txtSurname & txtForname = "" Then
1240                  i = InStr(1, txtName, " ")
1250                  If i = 0 Then
1260                      .ForeName = txtName
1270                      .SurName = ""
1280                  Else
1290                      .SurName = Left$(txtName, i - 1)
1300                      .ForeName = Mid$(txtName, i + 1, Len(txtName))
1310                  End If
1320              Else
1330                  .SurName = txtSurname
1340                  .ForeName = txtForname
1350              End If
1360              If tDoB <> "" Then
1370                  .DoB = tDoB
1380              End If
1390              .Sex = Left$(tSex, 1)
1400              .PatientGroup = tgroup
1410              .DeReservationDateTime = Format(DeReservationDateTime, "dd-MMM-yyyy hh:mm:ss")
1420              .ActionText = "Reserve Stock"
1430              .UserName = UserName
1440          End With
1450          LogCourierInterface "RS3", MSG



1460      Else
1470          UpdatePrintedLabels lLabNumber, _
                                  g.TextMatrix(RowNumbersToPrint(Y), 0), _
                                  g.TextMatrix(RowNumbersToPrint(Y), 2), _
                                  ProductBarCodeFor(g.TextMatrix(RowNumbersToPrint(Y), 3))
1480      End If

1490      UpdateStatusToXM ProductBarCodeFor(g.TextMatrix(RowNumbersToPrint(Y), 3)), _
                           g.TextMatrix(RowNumbersToPrint(Y), 0), _
                           g.TextMatrix(RowNumbersToPrint(Y), 2)

1500  Next
1510  cmdPDF.Caption = "Print PDF"

1520  UpdatePrinted lLabNumber, "Label"

1530  For Y = 1 To UBound(RowNumbersToPrint)
1540      frmCheckMatch.DisplayNumber = "Pack # " & Format$(Y)
1550      frmCheckMatch.Show 1
1560  Next

1570  FillDetails

1580  Exit Sub

cmdPDF_Click_Error:

      Dim strES As String
      Dim intEL As Integer

1590  intEL = Erl
1600  strES = Err.Description
1610  LogError "frmXMLabel", "cmdPDF_Click", intEL, strES, sql

End Sub
Private Sub UpdateStatusToXM(ByVal ProductCode As String, _
                             ByVal UnitNumber As String, _
                             ByVal Expiry As String)

          Dim Ps As New Products
          Dim p As Product

10    On Error GoTo UpdateStatusToXM_Error

20    Ps.LoadLatestISBT128 UnitNumber, ProductCode

30    If Ps.Count > 0 Then
40      Set p = Ps(1)
50      p.PackEvent = "X"
60      p.RecordDateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
70      p.EventEnd = ""
80      p.EventStart = ""
90      p.Chart = txtChart
100     p.PatName = txtName
110     p.UserName = UserCode

120     p.Save

130   End If

140   Exit Sub

UpdateStatusToXM_Error:

          Dim strES As String
          Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmXMLabel", "UpdateStatusToXM", intEL, strES

End Sub
Private Sub FillBatchHistory()

          Dim tb As Recordset
          Dim tb2 As Recordset
          Dim tb3 As Recordset
          Dim sql As String
          Dim s As String
          Dim BP As BatchProduct
          Dim BPs As New BatchProducts

10    On Error GoTo FillBatchHistory_Error

20    BPs.LoadSampleIDNoAudit mSampleID
30    For Each BP In BPs
40      If BP.EventCode = "I" Then
50          s = BP.BatchNumber & vbTab & _
                BP.UnitGroup & vbTab & _
                Format$(BP.DateExpiry, "dd/mm/yy") & vbTab & _
                BP.Product & vbTab & _
                BP.UserName & vbTab & _
                Format$(BP.RecordDateTime, "dd/mm/yyyy hh:mm:ss") & vbTab & _
                BP.Identifier
60          If BP.LabelPrinted Then
70              gPrinted.AddItem s
80          Else
90              g.AddItem s
100         End If
110     End If
120   Next

130   sql = "select distinct BatchNumber, Date from BatchDetails where " & _
          "SampleID = '" & mSampleID & "'"
140   Set tb = New Recordset
150   RecOpenClientBB 0, tb, sql

160   Do While Not tb.EOF
170     sql = "Select * from BatchDetails where " & _
              "BatchNumber = '" & tb!BatchNumber & "' " & _
              "and Date = '" & Format(tb!Date, "dd/mmm/yyyy hh:mm:ss") & "'"
180     Set tb2 = New Recordset
190     RecOpenClientBB 0, tb2, sql
200     s = tb2!BatchNumber & vbTab
210     sql = "Select * from BatchProductList where " & _
              "BatchNumber = '" & tb2!BatchNumber & "' " & _
              "and Product = '" & tb2!Product & "'"
220     Set tb3 = New Recordset
230     RecOpenClientBB 0, tb3, sql
240     If Not tb3.EOF Then
250         s = s & tb3!Group & vbTab & _
                Format(tb3!DateExpiry, "dd/mm/yy") & vbTab & _
                tb2!Product & vbTab & _
                TechnicianNameForCode(tb2!UserCode & "") & vbTab & _
                Format(tb2!Date, "dd/mm/yyyy hh:mm:ss")
260         g.AddItem s
270     End If
280     tb.MoveNext
290   Loop

300   Exit Sub

FillBatchHistory_Error:

          Dim strES As String
          Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "frmXMLabel", "FillBatchHistory", intEL, strES, sql


End Sub

Private Sub cmdReprintPDF_Click()

          Dim n As Integer
          Dim Y As Integer
          Dim HoldFor As String

10    On Error GoTo cmdReprintPDF_Click_Error

20    ReDim RowNumbersToPrint(1 To 1) As Integer

30    HoldFor = GetOptionSetting("TransfusionHoldFor", "72")

40    n = 0
50    gPrinted.col = 0
60    For Y = 1 To gPrinted.Rows - 1
70      gPrinted.row = Y
80      If gPrinted.CellBackColor = vbRed Then
90          n = n + 1
100         ReDim Preserve RowNumbersToPrint(1 To n) As Integer
110         RowNumbersToPrint(n) = Y
120     End If
130   Next

140   If n = 0 Then
150     iMsg "Nothing to do!" & vbCr & _
             "Click on units to re-print.", vbExclamation
160     If TimedOut Then Unload Me: Exit Sub
170     Exit Sub
180   End If

190   cmdReprintPDF.Caption = "Printing..."
200   For Y = 1 To UBound(RowNumbersToPrint)

210     PrintPDFCavan gPrinted.TextMatrix(RowNumbersToPrint(Y), 0), txtName, txtChart, tDoB, _
                      tward, tgroup, _
                      gPrinted.TextMatrix(RowNumbersToPrint(Y), 3), _
                      gPrinted.TextMatrix(RowNumbersToPrint(Y), 1), _
                      gPrinted.TextMatrix(RowNumbersToPrint(Y), 2), _
                      tcompat, _
                      tSampleDate, tSex, tgroup, txtSurname, txtForname

220   Next

230   cmdReprintPDF.Caption = "Print PDF"
240   For Y = 1 To UBound(RowNumbersToPrint)
250     frmCheckMatch.DisplayNumber = "Pack # " & Format$(Y)
260     frmCheckMatch.Show 1
270   Next

280   Exit Sub

cmdReprintPDF_Click_Error:

          Dim strES As String
          Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "frmXMLabel", "cmdReprintPDF_Click", intEL, strES

End Sub



Private Sub Form_Activate()

          Dim s As String

10    If Activated Then Exit Sub

20    Activated = True

30    s = "Have you removed units from Haemosafe Fridge?"
40    Answer = iMsg(s, vbQuestion + vbYesNo, , vbRed)
50    If TimedOut Then Unload Me: Exit Sub
60    If Answer <> vbYes Then
70      Unload Me
80      Exit Sub
90    End If

End Sub

Private Sub Form_Load()

          Dim Y As Integer
          Dim Product As String
          Dim Multi As Boolean

10    On Error GoTo Form_Load_Error

20    Activated = False

30    cmdPDF.Visible = True
40    bprint.Visible = False

50    FillDetails

60    Multi = False
70    g.col = 0
80    Product = g.TextMatrix(1, 3)
90    For Y = 1 To g.Rows - 1
100     If g.TextMatrix(Y, 3) = Product Then
110         g.row = Y
120         g.CellBackColor = vbRed
130         g.CellForeColor = vbYellow
140     Else
150         Multi = True
160     End If
170   Next

180   g.Sort = 9

190   If Multi Then
200     bSelect.Visible = False
210   End If

220   Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmXMLabel", "Form_Load", intEL, strES

End Sub
Private Sub FillDetails()

          Dim s As String
          Dim ok As Integer
          Dim issued As Integer
          Dim sql As String
          Dim tb As Recordset
          Dim Tx As Recordset

10    On Error GoTo FillDetails_Error

20    g.Font.Bold = True

30    g.Rows = 2
40    g.AddItem ""
50    g.RemoveItem 1
60    gPrinted.Rows = 2
70    gPrinted.AddItem ""
80    gPrinted.RemoveItem 1

90    sql = "Select * from PatientDetails where " & _
          "LabNumber = '" & mSampleID & "'"
100   Set tb = New Recordset
110   RecOpenServerBB 0, tb, sql
120   If Not tb.EOF Then
130     lTypenex = tb!Typenex & ""
140     tAandE = tb!AandE & ""
150     Select Case tb!Sex & ""
        Case "M": tSex = "Male"
160     Case "F": tSex = "Female"
170     Case Else: tSex = ""
180     End Select
190     tClinician = tb!Clinician & ""
200     txtName = tb!Name & ""
210     txtSurname = Trim$(tb!PatSurName & "")
220     txtForname = Trim$(tb!PatForeName & "")
230     tward = tb!Ward & ""
240     lOp = UserName
250     lAB = ""
260     If Trim$(tb!Anti3Reported & "") <> "" Then lAB = "Antibodies: " & tb!Anti3Reported
270     If Trim$(tb!AIDS & "") <> "" Then lAB = "Antibodies: " & tb!AIDS
280     If Trim$(tb!AIDR & "") <> "" Then lAB = "Antibodies: " & tb!AIDR
290     If lAB = "Antibodies: Negative" Then
300         lAB = "No Atypical Antibodies detected."
310     End If
320     tComment = StripComment(tb!Comment & "")
330     txtChart = tb!Patnum
340     txtAddr1 = tb!Addr1
350     txtAddr2 = tb!Addr2
360     tgroup = tb!fGroup & ""
370     tDoB = tb!DoB & ""
380     tSampleDate = Format$(tb!SampleDate, "dd/MM/yyyy HH:nn")
390     lLabNumber = mSampleID

400     sql = "SELECT DISTINCT L.ISBT128, L.DateExpiry, L.Event, L.PatName, L.GroupRh, L.BarCode, " & _
              "L.cRT, L.cRTr, L.cCO, L.cCOr, L.cEN, L.cENr, " & _
              "L.Operator, L.DateTime " & _
              "FROM Latest L, Product P WHERE " & _
              "L.ISBT128 = P.ISBT128 " & _
              "AND L.BarCode = P.BarCode " & _
              "AND L.DateExpiry = P.DateExpiry " & _
              "AND L.LabNumber = P.LabNumber " & _
              "AND L.LabNumber = '" & mSampleID & "' " & _
              "AND (L.Event = 'X' or " & _
            "     L.Event = 'I' or L.Event = 'K' or L.Event = 'V')"
410     Set Tx = New Recordset
420     RecOpenServerBB 0, Tx, sql
430     Do While Not Tx.EOF
440         ok = True: issued = False
450         If Tx!crtr Or Tx!ccor Or Tx!cenr Then ok = False
460         If Not ok Then
470             Answer = iMsg("Units are more compatible than the Patients Auto.", vbYesNo + vbQuestion)
480             If TimedOut Then Unload Me: Exit Sub
490             If Answer = vbYes Then
500                 ok = True
510             End If
520         End If
530         If Tx!Event = "I" Or Tx!Event = "V" Then
540             ok = True
550             issued = True
560         End If
570         If ok Or issued Then
580             s = Tx!ISBT128 & vbTab
590             s = s & Bar2Group(Tx!GroupRh) & vbTab
600             s = s & Format(Tx!DateExpiry, "dd/mmm/yyyy HH:mm") & vbTab
610             s = s & ProductWordingFor(Tx!BarCode) & vbTab
620             s = s & Tx!Operator & vbTab
630             s = s & Tx!DateTime
640             If AlreadyPrintedLabel(mSampleID, Tx!ISBT128, Format(Tx!DateExpiry, "dd/mmm/yyyy HH:mm"), Tx!BarCode & "") Then
650                 gPrinted.AddItem s
660             Else
670                 g.AddItem s
680             End If
690         End If
700         Tx.MoveNext
710     Loop
720   End If

730   FillBatchHistory

740   If g.Rows > 2 Then
750     g.RemoveItem 1
760   End If
770   If gPrinted.Rows > 2 Then
780     gPrinted.RemoveItem 1
790   End If

800   Exit Sub

FillDetails_Error:

          Dim strES As String
          Dim intEL As Integer

810   intEL = Erl
820   strES = Err.Description
830   LogError "frmXMLabel", "FillDetails", intEL, strES, sql

End Sub
Public Property Let SampleID(ByVal sNewValue As String)

10    mSampleID = sNewValue

End Property

Private Sub Form_Unload(Cancel As Integer)

10    Activated = False

End Sub

Private Sub g_Click()

          Dim Product As String
          Dim n As Integer

10    If g.MouseRow < 1 Then Exit Sub

20    g.col = 0
30    If g.CellBackColor = vbRed Then
40      g.CellBackColor = &H80000018
50      g.CellForeColor = &H8000000D
60    Else
70      Product = g.TextMatrix(g.row, 3)
80      g.CellBackColor = vbRed
90      g.CellForeColor = vbYellow
100     For n = 1 To g.Rows - 1
110         If g.TextMatrix(n, 3) <> Product Then
120             g.row = n
130             g.CellBackColor = &H80000018
140             g.CellForeColor = &H8000000D
150         End If
160     Next
170   End If

End Sub


Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

          Dim d1 As String
          Dim d2 As String

          'If Not IsDate(gXmatch.TextMatrix(Row1, 10)) Then
          '  Cmp = 0
          '  Exit Sub
          'End If
          '
          'If Not IsDate(gXmatch.TextMatrix(Row2, 10)) Then
          '  Cmp = 0
          '  Exit Sub
          'End If

10    d1 = Format(g.TextMatrix(Row1, 2), "dd/mmm/yyyy hh:mm:ss")
20    d2 = Format(g.TextMatrix(Row2, 2), "dd/mmm/yyyy hh:mm:ss")

30    Cmp = Sgn(DateDiff("D", d2, d1))

End Sub

Private Sub gPrinted_Click()

          Dim Product As String
          Dim n As Integer

10    With gPrinted
20      If .MouseRow < 1 Then Exit Sub
30      .col = 0
40      If .CellBackColor = vbRed Then
50          .CellBackColor = &H80000018
60          .CellForeColor = &H8000000D
70      Else
80          Product = .TextMatrix(.row, 3)
90          .CellBackColor = vbRed
100         .CellForeColor = vbYellow
110         For n = 1 To .Rows - 1
120             If .TextMatrix(n, 3) <> Product Then
130                 .row = n
140                 .CellBackColor = &H80000018
150                 .CellForeColor = &H8000000D
160             End If
170         Next
180     End If
190   End With

End Sub

Private Sub gPrinted_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

          Dim d1 As String
          Dim d2 As String

10    d1 = Format(gPrinted.TextMatrix(Row1, 2), "dd/mmm/yyyy hh:mm:ss")
20    d2 = Format(gPrinted.TextMatrix(Row2, 2), "dd/mmm/yyyy hh:mm:ss")
30    Cmp = Sgn(DateDiff("D", d2, d1))

End Sub


Private Sub Image1_Click()

10    If tcompat = "Electronically issued as compatible" Then
20        tcompat = "Is Group Compatible with Sample Labelled"
30    ElseIf tcompat = "Is Group Compatible with Sample Labelled" Then
40      tcompat = "Is Least Incompatible with Sample Labelled"
50    ElseIf tcompat = "Is Least Incompatible with Sample Labelled" Then
60      tcompat = "Is Issued for"
70    ElseIf tcompat = "Is Issued for" Then
80      tcompat = "Is Compatible with Serum From Sample Labelled"
90    ElseIf tcompat = "Is Compatible with Serum From Sample Labelled" Then
100     tcompat = "Is Compatible with"
110   ElseIf tcompat = "Is Compatible with" Then
120     tcompat = "Electronically issued as compatible"
130   End If

End Sub

Private Sub Image2_Click()

10    If tcompat = "Electronically issued as compatible" Then
20        tcompat = "Is Compatible with"
30    ElseIf tcompat = "Is Compatible with" Then
40      tcompat = "Is Compatible with Serum From Sample Labelled"
50    ElseIf tcompat = "Is Compatible with Serum From Sample Labelled" Then
60      tcompat = "Is Issued for"
70    ElseIf tcompat = "Is Issued for" Then
80      tcompat = "Is Least Incompatible with Sample Labelled"
90    ElseIf tcompat = "Is Least Incompatible with Sample Labelled" Then
100     tcompat = "Is Group Compatible with Sample Labelled"
110   ElseIf tcompat = "Is Group Compatible with Sample Labelled" Then
120     tcompat = "Electronically issued as compatible"
130   End If

End Sub




Private Sub tmrCourier_Timer()

          Dim C As Control

10    On Error GoTo tmrCourier_Timer_Error

20    pbCourier.Value = pbCourier.Value + 1
30    If pbCourier.Value = pbCourier.max Then

40      For Each C In Me.Controls
50          C.Enabled = True
60      Next

70      fraCourier.Visible = False
80      pbCourier.Value = 0
90      tmrCourier.Enabled = False

100   End If

110   Exit Sub

tmrCourier_Timer_Error:

          Dim strES As String
          Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmXMLabel", "tmrCourier_Timer", intEL, strES

End Sub


