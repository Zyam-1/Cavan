VERSION 5.00
Begin VB.Form frmMaintenance 
   Caption         =   "NetAcquire"
   ClientHeight    =   6330
   ClientLeft      =   3690
   ClientTop       =   2400
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   6150
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   2490
      Picture         =   "frmMaintenance.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5340
      Width           =   1245
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove All References"
      Height          =   615
      Left            =   3060
      Picture         =   "frmMaintenance.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1020
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto Select From"
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   150
      Width           =   1845
      Begin VB.OptionButton optAuto 
         Caption         =   "Demographics"
         Height          =   255
         Index           =   4
         Left            =   210
         TabIndex        =   7
         Top             =   1500
         Width           =   1365
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Externals"
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   6
         Top             =   1200
         Width           =   1005
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Coagulation"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   5
         Top             =   900
         Width           =   1185
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Biochemistry"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   4
         Top             =   600
         Width           =   1305
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Haematology"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.TextBox txtSampleID 
      Height          =   285
      Left            =   3060
      TabIndex        =   0
      Top             =   660
      Width           =   1845
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMaintenance.frx":0CD4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2505
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   5685
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SampleID"
      Height          =   195
      Left            =   3600
      TabIndex        =   1
      Top             =   450
      Width           =   690
   End
End
Attribute VB_Name = "frmMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AutoFill(ByVal Index As Integer)

      Dim tb As Recordset
      Dim sql As String

34590 On Error GoTo AutoFill_Error

34600 sql = "Select top 1 SampleID from "

34610 Select Case Index
        Case 0: sql = sql & "HaemResults"
34620   Case 1: sql = sql & "BioResults"
34630   Case 2: sql = sql & "CoagResults"
34640   Case 3: sql = sql & "ExtResults"
34650   Case 4: sql = sql & "Demographics"
34660 End Select

34670 sql = sql & " order by SampleID Desc"
34680 Set tb = New Recordset
34690 RecOpenServer 0, tb, sql
34700 If Not tb.EOF Then
34710   txtSampleID = tb!SampleID
34720 Else
34730   txtSampleID = ""
34740 End If

34750 Exit Sub

AutoFill_Error:

      Dim strES As String
      Dim intEL As Integer

34760 intEL = Erl
34770 strES = Err.Description
34780 LogError "frmMaintenance", "AutoFill", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

34790 Unload Me

End Sub

Private Sub cmdDelete_Click()

      Dim n As Integer
      Dim sql As String

34800 On Error GoTo cmdDelete_Click_Error

34810 If iMsg("Are you sure?", vbQuestion + vbYesNo, , vbRed, 18) = vbYes Then

34820   sql = "Delete from HaemResults where " & _
              "SampleID = '" & txtSampleID & "'"
34830   Cnxn(0).Execute sql

34840   sql = "Delete from BioResults where " & _
              "SampleID = '" & txtSampleID & "'"
34850   Cnxn(0).Execute sql

34860   sql = "Delete from CoagResults where " & _
              "SampleID = '" & txtSampleID & "'"
34870   Cnxn(0).Execute sql

34880   sql = "Delete from ExtResults where " & _
              "SampleID = '" & txtSampleID & "'"
34890   Cnxn(0).Execute sql

34900   sql = "Delete from Demographics where " & _
              "SampleID = '" & txtSampleID & "'"
34910   Cnxn(0).Execute sql

34920 End If

34930 For n = 0 To 4
34940   If optAuto(n) Then
34950     AutoFill n
34960     Exit For
34970   End If
34980 Next

34990 Exit Sub

cmdDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

35000 intEL = Erl
35010 strES = Err.Description
35020 LogError "frmMaintenance", "cmdDelete_Click", intEL, strES, sql


End Sub


Private Sub optAuto_Click(Index As Integer)

35030 AutoFill Index

End Sub


