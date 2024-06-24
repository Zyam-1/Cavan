VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewProducts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Request Details"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCL 
      Height          =   1185
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   5190
      Width           =   5595
   End
   Begin VB.TextBox txtNotes 
      Height          =   1185
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3690
      Width           =   5595
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   465
      Left            =   6780
      TabIndex        =   1
      Top             =   4410
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid flxDetail 
      Height          =   3195
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5636
      _Version        =   393216
      Cols            =   5
   End
   Begin VB.Label Label2 
      Caption         =   "Clinical Detail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   3420
      Width           =   1665
   End
End
Attribute VB_Name = "frmViewProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public g_SampleID As String

Private Const fcsLine_NO = 0
Private Const fcsQes = 1
Private Const fcsAns = 2
Private Const fcsRID = 3
Private Const fcsSID = 4

Private Sub FormatGrid()
    On Error GoTo ERROR_FormatGrid
    
    flxDetail.Rows = 1
    flxDetail.row = 0
    
    flxDetail.ColWidth(fcsLine_NO) = 230
    
    flxDetail.TextMatrix(0, fcsQes) = "Questions"
    flxDetail.ColWidth(fcsQes) = 3350
    flxDetail.ColAlignment(fcsQes) = flexAlignLeftCenter
    
    flxDetail.TextMatrix(0, fcsAns) = "Answers"
    flxDetail.ColWidth(fcsAns) = 4000
    flxDetail.ColAlignment(fcsAns) = flexAlignLeftCenter
    
    flxDetail.TextMatrix(0, fcsRID) = ""
    flxDetail.ColWidth(fcsRID) = 0
    flxDetail.ColAlignment(fcsRID) = flexAlignLeftCenter
    
    flxDetail.TextMatrix(0, fcsSID) = ""
    flxDetail.ColWidth(fcsSID) = 0
    flxDetail.ColAlignment(fcsSID) = flexAlignLeftCenter
    
        
    Exit Sub
ERROR_FormatGrid:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmViewProducts", "FormatGrid", intEL, strES
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub




Private Sub Form_Load()
On Error GoTo ERROR_Form_Load
     'Zyam commented subs that are related to ocm 26-1-24
'    Call FormatGrid
'    Call ShowDetail
     'Zyam

    Exit Sub
ERROR_Form_Load:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmViewProducts", "Form_Load", intEL, strES
End Sub

Public Sub ShowDetail()
    On Error GoTo ERROR_ShowDetail
    
    Dim sql As String
    Dim tb As ADODB.Recordset
    Dim tbQ As ADODB.Recordset
    Dim tbR As ADODB.Recordset
    Dim l_str As String
    
    flxDetail.Rows = 1
    flxDetail.row = 0
    
    txtNotes.Text = ""
    
    sql = "Select RequestID from OcmRequestDetails Where SampleID = '" & g_SampleID & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If Not tb Is Nothing Then
        If Not tb.EOF Then
            sql = "Select question,answer from ocmQuestions Where rid = '" & tb!RequestID & "'"
            Set tbQ = New Recordset
            RecOpenServer 0, tbQ, sql
            If Not tbQ Is Nothing Then
                If Not tbQ.EOF Then
                    While Not tbQ.EOF
                        l_str = "" & vbTab & tbQ!question & vbTab & tbQ!Answer & vbTab & tb!RequestID & vbTab & g_SampleID
                        flxDetail.AddItem (l_str)
                        tbQ.MoveNext
                    Wend
                End If
            End If
            sql = "Select Notes, CLDetails from ocmRequest Where RequestID = '" & tb!RequestID & "'"
            Set tbR = New Recordset
            RecOpenServer 0, tbR, sql
            If Not tbR Is Nothing Then
                If Not tbR.EOF Then
                    txtNotes.Text = tbR!Notes
                    txtCL.Text = tbR!CLDetails
                End If
            End If
        End If
    End If
    
    
        
    Exit Sub
ERROR_ShowDetail:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmViewProducts", "ShowDetail", intEL, strES
End Sub

