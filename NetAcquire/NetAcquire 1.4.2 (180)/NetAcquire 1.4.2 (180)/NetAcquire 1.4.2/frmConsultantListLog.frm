VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultantListLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   5700
      Picture         =   "frmConsultantListLog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2985
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5265
      _Version        =   393216
      Rows            =   4
      FixedRows       =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
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
End
Attribute VB_Name = "frmConsultantListLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_sSampleID As String

Private Sub bcancel_Click()

23890     On Error GoTo bCancel_Click_Error

23900     Unload Me

23910     Exit Sub

bCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

23920     intEL = Erl
23930     strES = Err.Description
23940     LogError "frmConsultantListLog", "bCancel_Click", intEL, strES
          
End Sub

Private Sub Form_Load()

23950     On Error GoTo Form_Load_Error


23960     With Me
23970         .Caption = "Consultant List Log"
23980     End With

23990     LoadConsultantListLog (SampleID)
24000     Exit Sub


Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

24010     intEL = Erl
24020     strES = Err.Description
24030     LogError "frmConsultantListLog", "Form_Load", intEL, strES
End Sub

Public Property Get SampleID() As String

24040     SampleID = m_sSampleID

End Property

Public Property Let SampleID(ByVal sSampleID As String)

24050     m_sSampleID = sSampleID

End Property


Private Sub GridHead()

24060     On Error GoTo GridHead_Error


24070     With g
24080         .Clear

24090         .Rows = 1
24100         .FixedCols = 0
24110         .Cols = 3

24120         .TextMatrix(0, 0) = "Dated"
24130         .ColWidth(0) = 2100
24140         .ColAlignment(0) = 1

24150         .TextMatrix(0, 1) = "User Name"
24160         .ColWidth(1) = 2000
24170         .ColAlignment(1) = 1

24180         .TextMatrix(0, 2) = "Status"
24190         .ColWidth(2) = 2400
24200         .ColAlignment(2) = 1
24210     End With


24220     Exit Sub


GridHead_Error:

          Dim strES As String
          Dim intEL As Integer

24230     intEL = Erl
24240     strES = Err.Description
24250     LogError "frmConsultantListLog", "GridHead", intEL, strES
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadConsultantListLog
' Author    : Masood
' Date      : 28/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
 Sub LoadConsultantListLog(SampleID As String)

24260     On Error GoTo LoadConsultantListLog_Error

          Dim sql As String
          Dim tb As ADODB.Recordset
          Dim s As String
24270     GridHead
24280     sql = "Select  SampleID, UserName, Status, DateTimeOfRecord FROM ConsultantListLog where SampleID = '" & SampleID & "' Order by DateTimeOfRecord Desc"
24290     Set tb = New Recordset

24300     RecOpenClient 0, tb, sql

24310     Do While Not tb.EOF
24320         s = tb!DateTimeOfRecord & vbTab & tb!UserName & vbTab & tb!Status
24330         g.AddItem s
24340         tb.MoveNext
24350     Loop



24360     Exit Sub


LoadConsultantListLog_Error:

          Dim strES As String
          Dim intEL As Integer

24370     intEL = Erl
24380     strES = Err.Description
24390     LogError "frmConsultantListLog", "LoadConsultantListLog", intEL, strES, sql
End Sub

