VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAntibodyIdent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Antibody Identification"
   ClientHeight    =   7755
   ClientLeft      =   285
   ClientTop       =   450
   ClientWidth     =   9855
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
   ForeColor       =   &H00808080&
   Icon            =   "frmAntibodyIdent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7755
   ScaleWidth      =   9855
   Begin VB.OptionButton optReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6810
      TabIndex        =   53
      Top             =   6930
      Width           =   855
   End
   Begin VB.OptionButton optReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6810
      TabIndex        =   52
      Top             =   5040
      Width           =   855
   End
   Begin VB.OptionButton optReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6810
      TabIndex        =   51
      Top             =   3210
      Width           =   855
   End
   Begin VB.OptionButton optReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   6780
      TabIndex        =   50
      Top             =   1320
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   8100
      Picture         =   "frmAntibodyIdent.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6300
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   765
      Left            =   8100
      Picture         =   "frmAntibodyIdent.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5430
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   3
      Left            =   60
      TabIndex        =   33
      Top             =   5670
      Width           =   6705
      Begin VB.ComboBox cID 
         Height          =   315
         Index           =   3
         Left            =   630
         TabIndex        =   46
         Text            =   "cID"
         Top             =   180
         Width           =   2265
      End
      Begin VB.TextBox txtReport 
         BackColor       =   &H80000014&
         Height          =   585
         Index           =   3
         Left            =   3600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   1080
         Width           =   2955
      End
      Begin VB.CommandButton bview 
         Appearance      =   0  'Flat
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2940
         TabIndex        =   34
         Top             =   210
         Width           =   675
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   405
         Index           =   3
         Left            =   180
         TabIndex        =   36
         Top             =   510
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   714
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   2
         ScrollBars      =   0
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
      Begin VB.Label lCoEn 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coombs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5610
         TabIndex        =   42
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lTemp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "37 deg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4860
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Panel"
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
         Index           =   3
         Left            =   180
         TabIndex        =   40
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Suggest"
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
         Index           =   3
         Left            =   180
         TabIndex        =   39
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lSuggest 
         BorderStyle     =   1  'Fixed Single
         Height          =   585
         Index           =   3
         Left            =   900
         TabIndex        =   38
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Report"
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
         Index           =   3
         Left            =   3000
         TabIndex        =   37
         Top             =   1140
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   2
      Left            =   60
      TabIndex        =   23
      Top             =   3810
      Width           =   6705
      Begin VB.ComboBox cID 
         Height          =   315
         Index           =   2
         Left            =   660
         TabIndex        =   45
         Text            =   "cID"
         Top             =   180
         Width           =   2265
      End
      Begin VB.CommandButton bview 
         Appearance      =   0  'Flat
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2940
         TabIndex        =   26
         Top             =   210
         Width           =   675
      End
      Begin VB.TextBox txtReport 
         BackColor       =   &H80000014&
         Height          =   585
         Index           =   2
         Left            =   3600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   1080
         Width           =   2955
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   405
         Index           =   2
         Left            =   180
         TabIndex        =   24
         Top             =   510
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   714
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   2
         ScrollBars      =   0
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Report"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   32
         Top             =   1140
         Width           =   480
      End
      Begin VB.Label lSuggest 
         BorderStyle     =   1  'Fixed Single
         Height          =   585
         Index           =   2
         Left            =   900
         TabIndex        =   31
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Suggest"
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
         Index           =   2
         Left            =   180
         TabIndex        =   30
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Panel"
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
         Index           =   2
         Left            =   180
         TabIndex        =   29
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lTemp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "37 deg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4860
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lCoEn 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coombs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5610
         TabIndex        =   27
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   1
      Left            =   60
      TabIndex        =   13
      Top             =   1950
      Width           =   6705
      Begin VB.ComboBox cID 
         Height          =   315
         Index           =   1
         Left            =   630
         TabIndex        =   44
         Text            =   "cID"
         Top             =   180
         Width           =   2325
      End
      Begin VB.CommandButton bview 
         Appearance      =   0  'Flat
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2970
         TabIndex        =   16
         Top             =   180
         Width           =   675
      End
      Begin VB.TextBox txtReport 
         BackColor       =   &H80000014&
         Height          =   585
         Index           =   1
         Left            =   3600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1080
         Width           =   2955
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   405
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   510
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   714
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   2
         ScrollBars      =   0
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Report"
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
         Left            =   3000
         TabIndex        =   22
         Top             =   1140
         Width           =   480
      End
      Begin VB.Label lSuggest 
         BorderStyle     =   1  'Fixed Single
         Height          =   585
         Index           =   1
         Left            =   900
         TabIndex        =   21
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Suggest"
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
         Left            =   180
         TabIndex        =   20
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Panel"
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
         Left            =   180
         TabIndex        =   19
         Top             =   210
         Width           =   405
      End
      Begin VB.Label lTemp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "37 deg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4860
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lCoEn 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coombs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5610
         TabIndex        =   17
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   6705
      Begin VB.ComboBox cID 
         Height          =   315
         Index           =   0
         Left            =   660
         TabIndex        =   43
         Text            =   "cID"
         Top             =   180
         Width           =   2295
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   405
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   510
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   714
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   2
         ScrollBars      =   0
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
      Begin VB.TextBox txtReport 
         BackColor       =   &H80000014&
         Height          =   585
         Index           =   0
         Left            =   3600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1080
         Width           =   2955
      End
      Begin VB.CommandButton bview 
         Appearance      =   0  'Flat
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2970
         TabIndex        =   4
         Top             =   210
         Width           =   675
      End
      Begin VB.Label lCoEn 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coombs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5610
         TabIndex        =   12
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lTemp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "37 deg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4860
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Panel"
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
         Left            =   210
         TabIndex        =   9
         Top             =   210
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Suggest"
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
         Left            =   180
         TabIndex        =   8
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lSuggest 
         BorderStyle     =   1  'Fixed Single
         Height          =   585
         Index           =   0
         Left            =   900
         TabIndex        =   7
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Report"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   1140
         Width           =   480
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   660
      TabIndex        =   54
      Top             =   7530
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
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
      Left            =   7050
      TabIndex        =   49
      Top             =   570
      Width           =   735
   End
   Begin VB.Label lblSampleID 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7020
      TabIndex        =   2
      Top             =   780
      Width           =   2595
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Patient Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7080
      TabIndex        =   0
      Top             =   30
      Width           =   960
   End
   Begin VB.Label lname 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7020
      TabIndex        =   1
      Top             =   240
      Width           =   2595
   End
End
Attribute VB_Name = "frmAntibodyIdent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub FillDetails()

      Dim n As Integer
      Dim tb As Recordset
      Dim sql As String
      Dim X As Long
      Dim reaction As String

10    On Error GoTo FillDetails_Error

20    For n = 0 To 3
30      sql = "Select * from ABResults where " & _
              "[Index] = '" & n & "' " & _
              "and SampleID = '" & lblSampleID & "'"
40      Set tb = New Recordset
50      RecOpenServerBB 0, tb, sql
60      If Not tb.EOF Then
70        cID(n).Text = tb!PanelNumber & ""
          'tb!Operator = Left$(UserCode, 4)
          'tb!DateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
80        g(n).Col = 0
90        reaction = tb!reaction & ""
100       If Len(reaction) >= 2 Then
110         g(n).Cols = Len(reaction) \ 2 + 1
120         For X = 0 To g(n).Cols - 1
130           g(n).ColWidth(X) = 300
140         Next
150         g(n).ColSel = g(n).Cols - 1
160       End If
170       g(n).Clip = reaction
180       lTemp(n) = tb!temperature & ""
190       lCoEn(n) = tb!CorE & ""
          'tb!Index = n
          'tb!SampleID = lblSampleID
200       optReport(n) = IIf(tb!Reported, True, False)
210       lSuggest(n) = tb!Suggest & ""
220       txtReport(n) = tb!Report & ""
230       tb.Update
240     End If
250   Next

260   Exit Sub

FillDetails_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmAntibodyIdent", "FillDetails", intEL, strES, sql


End Sub

Private Sub LoadPanel(Index As Integer)

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim Y As Integer
      Dim X As Integer
      Dim Pattern As String

10    On Error GoTo LoadPanel_Error

20    g(Index).Cols = 50

30    s = Trim$(cID(Index))
40    If s = "" Then
50      iMsg "Specify Panel Number.", vbExclamation
60      If TimedOut Then Unload Me: Exit Sub
70      Exit Sub
80    End If

90    sql = "Select * from AntibodyPanels where " & _
            "LotNumber = '" & cID(Index) & "'"

100   Set tb = New Recordset
110   RecOpenServerBB 0, tb, sql

120   If tb.EOF Then
130     iMsg "Panel Number not found.", vbInformation
140     If TimedOut Then Unload Me: Exit Sub
150     Exit Sub
160   End If

170   If DateDiff("d", Now, tb!ExpiryDate) < 0 Then
180     iMsg "Panel Number has expired!", vbCritical
190     If TimedOut Then Unload Me: Exit Sub
200     Exit Sub
210   End If

220   sql = "Select top 1 * from AntibodyPatterns where " & _
            "LotNumber = '" & cID(Index) & "' " & _
            "and Position > 3"
230   Set tb = New Recordset
240   RecOpenServerBB 0, tb, sql
250   If tb.EOF Then Exit Sub

260   Y = 0
270   For X = 1 To Len(tb!Pattern)
280     If Mid$(tb!Pattern, X, 1) = vbTab Then
290       Y = Y + 1
300     End If
310   Next
320   If Right$(tb!Pattern, 1) = vbTab Then
330     Y = Y - 1
340   End If
350   g(Index).Cols = Y
360   For X = 0 To Y - 1
370     g(Index).ColWidth(X) = TextWidth("WW")
380     g(Index).ColAlignment(X) = flexAlignCenterCenter
390   Next

400   Exit Sub

LoadPanel_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "frmAntibodyIdent", "LoadPanel", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdSave_Click()

      Dim n As Integer
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    For n = 0 To 3
30      If cID(n).Text <> "" Then
40        sql = "Select * from ABResults where " & _
                "[Index] = '" & n & "' " & _
                "and SampleID = '" & lblSampleID & "'"
50        Set tb = New Recordset
60        RecOpenServerBB 0, tb, sql
70        If tb.EOF Then
80          tb.AddNew
90        End If
100       tb!PanelNumber = cID(n).Text
110       tb!Operator = Left$(UserCode, 4)
120       tb!DateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
130       g(n).Col = 0
140       g(n).ColSel = g(n).Cols - 1
150       tb!reaction = g(n).Clip
160       tb!temperature = lTemp(n)
170       tb!CorE = lCoEn(n)
180       tb!Index = n
190       tb!SampleID = lblSampleID
200       tb!Reported = optReport(n)
210       tb!Suggest = lSuggest(n)
220       tb!Report = txtReport(n)
230       tb.Update
240     End If
  
250     If optReport(n) Then
260       frmxmatch.tident = txtReport(n)
270     End If
280   Next

290   Unload Me

300   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "frmAntibodyIdent", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub bview_Click(Index As Integer)

10    With frmDefineABPanel
20      .Panel = cID(Index)
30      .Show 1
40    End With

End Sub

Private Sub cID_Click(Index As Integer)

10    LoadPanel (Index)

End Sub




Private Sub Form_Activate()

10    FillDetails

End Sub

Private Sub Form_Load()

      Dim sn As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo Form_Load_Error

20    sql = "Select LotNumber from AntibodyPanels " & _
            "where ExpiryDate >= '" & Format(Now, "dd/mmm/yyyy") & "' " & _
            "order by DateEntered desc"
30    Set sn = New Recordset
40    RecOpenServerBB 0, sn, sql

50    For n = 0 To 3
60      cID(n).Clear
70    Next

80    Do While Not sn.EOF
90      For n = 0 To 3
100       cID(n).AddItem sn!LotNumber & ""
110     Next
120     sn.MoveNext
130   Loop

140   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmAntibodyIdent", "Form_Load", intEL, strES, sql


End Sub
Private Sub g_Click(Index As Integer)

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim n As Integer
      Dim Found As Boolean
      Dim Possible As Boolean
      Dim IsFilled As Boolean
      Dim Pattern As String
      Dim PosPattern As String

10    On Error GoTo g_Click_Error

20    lSuggest(Index) = ""

30    If Trim$(cID(Index)) = "" Then
40      Exit Sub
50    End If

60    s = g(Index)
70    Select Case s
        Case "", " ": s = "O"
80      Case "O": s = "1"
90      Case "1": s = "2"
100     Case "2": s = "3"
110     Case "3": s = "4"
120     Case "4": s = "+"
130     Case "+": s = ""
140   End Select
150   g(Index) = s

160   IsFilled = True
170   For n = 0 To g(Index).Cols - 1
180     If Trim$(g(Index).TextMatrix(0, n)) = "" Then IsFilled = False
190   Next
200   If Not IsFilled Then
210     lSuggest(Index) = ""
220     Exit Sub
230   End If

240   Pattern = ""
250   For n = 0 To g(Index).Cols - 1
260     Pattern = Pattern & g(Index).TextMatrix(0, n) & vbTab
270   Next

280   If Right$(Pattern, 1) = vbTab Then
290     Pattern = Left$(Pattern, Len(Pattern) - 1)
300   End If

310   Found = False
320   Possible = False

330   PosPattern = Replace(Pattern, "1", "+")
340   PosPattern = Replace(PosPattern, "2", "+")
350   PosPattern = Replace(PosPattern, "3", "+")
360   PosPattern = Replace(PosPattern, "4", "+")

370   sql = "Select * from AntibodyPatterns where " & _
            "LotNumber = '" & cID(Index) & "' " & _
            "and Pattern like '%" & PosPattern & "'"
380   Set tb = New Recordset
390   RecOpenServerBB 0, tb, sql
400   If tb.EOF Then
410     PosPattern = Replace(PosPattern, "+", "_")
420     sql = "Select * from AntibodyPatterns where " & _
              "LotNumber = '" & cID(Index) & "' " & _
              "and Pattern like '%" & PosPattern & "'"
430     Set tb = New Recordset
440     RecOpenServerBB 0, tb, sql
450     If Not tb.EOF Then
460       Possible = True
470     End If
480   Else
490     Found = True
500   End If
510   Do While Not tb.EOF
520     n = InStr(tb!Pattern, vbTab)
530     lSuggest(Index) = lSuggest(Index) & Left$(tb!Pattern, n - 1) & " "
540     tb.MoveNext
550   Loop
  
560   If Possible Then
570     lSuggest(Index) = "Best Fit: " & lSuggest(Index)
580   Else
590     If Not Found Then
600       lSuggest(Index) = "Unable to identify."
610     End If
620   End If

630   Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

640   intEL = Erl
650   strES = Err.Description
660   LogError "frmAntibodyIdent", "g_Click", intEL, strES, sql


End Sub

Private Sub lsuggest_Change(Index As Integer)

10    txtReport(Index) = lSuggest(Index).Caption

End Sub

Private Sub lCoEn_Click(Index As Integer)

      Dim s As String

10    s = Trim$(lCoEn(Index))

20    Select Case s
        Case "Coombs": s = "Enzyme"
30      Case "Enzyme": s = "Saline"
40      Case "Saline": s = "Enz/Coombs"
50      Case "Enz/Coombs": s = "Coombs"
60    End Select

70    lCoEn(Index) = s

End Sub

Private Sub lTemp_Click(Index As Integer)

      Dim s As String

10    s = Trim$(lTemp(Index))

20    Select Case s
        Case "4 deg": s = "37 deg"
30      Case "Room": s = "4 deg"
40      Case "37 deg": s = "Room"
50    End Select

60    lTemp(Index) = s

End Sub

