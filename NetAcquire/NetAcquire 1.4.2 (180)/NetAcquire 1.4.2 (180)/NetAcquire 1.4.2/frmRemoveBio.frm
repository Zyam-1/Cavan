VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRemoveBio 
   AutoRedraw      =   -1  'True
   Caption         =   "NetAcquire"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewArchive 
      Caption         =   "No Archived Results"
      Height          =   525
      Left            =   4410
      TabIndex        =   5
      Top             =   1020
      Width           =   2085
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2715
      Left            =   330
      TabIndex        =   4
      Top             =   1890
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4789
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "Retrieve Analyte from Archive"
      Height          =   525
      Left            =   5490
      TabIndex        =   3
      Top             =   4950
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   7200
      Picture         =   "frmRemoveBio.frx":0000
      TabIndex        =   2
      Top             =   1020
      Width           =   1065
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Remove All Results"
      Height          =   525
      Left            =   360
      TabIndex        =   1
      Top             =   1020
      Width           =   3255
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Analyte from Results"
      Height          =   525
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5190
      TabIndex        =   8
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   4410
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   570
      Picture         =   "frmRemoveBio.frx":066A
      Top             =   4650
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Analyte(s) to be retrieved"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   5100
      Width           =   2610
   End
End
Attribute VB_Name = "frmRemoveBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pAnalyte As String

Private pSampleID As Long

Public Property Get Analyte() As String

Analyte = pAnalyte

End Property

Public Property Let Analyte(ByVal strNewValue As String)

pAnalyte = strNewValue

End Property
Public Property Let SampleID(ByVal lngNewValue As Long)

pSampleID = lngNewValue

End Property

