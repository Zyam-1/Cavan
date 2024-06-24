VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAVupdate 
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1065
      Left            =   2400
      TabIndex        =   2
      Top             =   630
      Width           =   2445
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   675
         TabIndex        =   3
         Top             =   270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         _Version        =   393216
         Format          =   127991809
         CurrentDate     =   43101
         MinDate         =   43101
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   675
         TabIndex        =   4
         Top             =   630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         _Version        =   393216
         Format          =   127991809
         CurrentDate     =   38631
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   6
         Top             =   315
         Width           =   345
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   675
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1275
      Left            =   7080
      Picture         =   "frmAVupdate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4245
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   1275
      Left            =   7080
      Picture         =   "frmAVupdate.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1950
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmAVupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
Dim Sql As String
Dim Sql2 As String
Dim tb As Recordset
Dim Iso As Recordset
Dim strSignOffDateTime As String

Sql = "SELECT * From PrintValidLog where (signoff = 0 or signoff is null) and Department = 'M' and valid = 1 and printed = 1 "
Sql = Sql & "and (ValidatedDateTime > '" & FromDate & "'  and ValidatedDateTime < '" & ToDate & "')"

Set tb = New Recordset
RecOpenClient Cn, tb, Sql

Do While Not Cn.EOF
    strSignOffDateTime = Format(tb!ValidatedDateTime, "dd/mmm/yyyy hh:mm:ss")
    Sql2 = "SELECT * From Isolates WHERE  sampleid = '" & tb!SampleID & "' and (OrganismGroup = 'Negative Results')"
    Set Iso = New Recordset
    RecOpenClient Cn, Iso, Sql2
    If Not Iso.EOF Then
        ' SignOff, SignOffBy, SignOffDateTime
        Sql = "Update PrintValidLog " & _
            "set SignOff = 1, SignOffBy = 'AV' where " & _
            "sampleid = '" & SampleID & "'"
        Cnxn(0).Execute Sql
    
    End If

    Cn.MoveNext
Loop

End Sub

Private Sub Form_Load()

End Sub
