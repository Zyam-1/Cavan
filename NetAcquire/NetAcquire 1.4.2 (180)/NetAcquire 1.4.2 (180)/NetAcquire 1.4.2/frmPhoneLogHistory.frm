VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPhoneLogHistory 
   Caption         =   "NetAcquire - Phone Log"
   ClientHeight    =   5100
   ClientLeft      =   75
   ClientTop       =   480
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   12090
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   345
      Left            =   4500
      TabIndex        =   4
      Top             =   180
      Width           =   975
   End
   Begin VB.TextBox txtSampleID 
      Height          =   315
      Left            =   2910
      TabIndex        =   3
      Top             =   180
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   10410
      Picture         =   "frmPhoneLogHistory.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4290
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid grdPhoneLog 
      Height          =   3585
      Left            =   270
      TabIndex        =   1
      Top             =   600
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   6324
      _Version        =   393216
      Cols            =   14
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmPhoneLogHistory.frx":066A
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   8550
      Picture         =   "frmPhoneLogHistory.frx":0725
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Phone Log History for Sample ID"
      Height          =   195
      Left            =   540
      TabIndex        =   0
      Top             =   240
      Width           =   2310
   End
End
Attribute VB_Name = "frmPhoneLogHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String

Private pPhoneOrFAX As String
Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim strRecord As String
      Dim intN As Integer
      Dim TableName As String
      Dim RootName As String

36870 On Error GoTo ehFG

36880 Screen.MousePointer = vbHourglass

36890 RootName = IIf(pPhoneOrFAX = "FAX", "Fax", "Phon")
36900 TableName = IIf(pPhoneOrFAX = "FAX", "FaxLog", "PhoneLog")

36910 With grdPhoneLog
36920   .Rows = 2
36930   .AddItem ""
36940   .RemoveItem 1
36950   .Visible = False
        
36960   sql = "SELECT * FROM " & TableName & " WHERE " & _
              "SampleID = " & Val(txtSampleID) & " " & _
              "ORDER BY DateTime DESC"

36970   Set tb = New Recordset
36980   RecOpenServer 0, tb, sql
36990   Do While Not tb.EOF
37000     strRecord = Format$(tb!DateTime, "dd/mm/yy hh:mm")
37010     For intN = 0 To 9
37020       strRecord = strRecord & vbTab
37030     Next
37040     strRecord = strRecord & tb(RootName & "edTo") & vbTab & _
                      tb!Comment & vbTab & _
                      tb(RootName & "edBy") & vbTab
37050     If RootName = "Phon" Then
37060       strRecord = strRecord & tb!Direction & ""
37070     End If
          
37080     .AddItem strRecord
        
37090     .row = .Rows - 1
37100     For intN = 1 To 9
37110       If InStr(tb!Discipline, Mid$("HBCIGEMDN", intN, 1)) Then
37120         .Col = intN
37130         Set .CellPicture = imgSquareTick.Picture
37140         .CellPictureAlignment = flexAlignCenterCenter
37150       End If
37160     Next
        
37170     tb.MoveNext

37180   Loop

37190   If .Rows > 2 Then
37200     .RemoveItem 1
37210   End If
37220   .Visible = True
37230 End With

37240 Screen.MousePointer = 0

37250 Exit Sub

ehFG:

37260 grdPhoneLog.Visible = True
37270 Screen.MousePointer = 0

37280 Exit Sub

End Sub

Private Sub cmdCancel_Click()

37290 Unload Me

End Sub


Private Sub cmdSearch_Click()

37300 If Trim$(txtSampleID) = "" Then Exit Sub

37310 FillG

End Sub

Private Sub Form_Activate()

37320 txtSampleID = pSampleID

37330 FillG

End Sub

Public Property Let SampleID(ByVal strNewValue As String)

37340 pSampleID = strNewValue

End Property


Private Sub Form_Load()

37350 If pPhoneOrFAX = "FAX" Then
37360   Label1.Caption = "FAX Log History for Sample ID"
37370   Me.Caption = "NetAcquire - FAX Log"
37380   grdPhoneLog.FormatString = "<Date Time                   |" & _
                                   "^Hae |^Bio |^Coa |^Imm |^Gas |^Ext |^Mic |^End |<Not Phoned |" & _
                                   "<Faxed To                |" & _
                                   "<Comment                                      |" & _
                                   "<Faxed By "
37390 Else
37400   Label1.Caption = "Phone Log History for Sample ID"
37410   Me.Caption = "NetAcquire - Phone Log"
37420   grdPhoneLog.FormatString = "<Date Time                   |" & _
                                   "^Hae |^Bio |^Coa |^Imm |^Gas |^Ext |^Mic |^End |<Not Phoned |" & _
                                   "<Phoned To            |" & _
                                   "<Comment                                      |" & _
                                   "<Phoned By "
37430 End If

End Sub

Private Sub grdPhoneLog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim s As String

37440 If grdPhoneLog.MouseCol = 0 Or grdPhoneLog.MouseCol > 8 Or grdPhoneLog.MouseRow = 0 Then
37450   grdPhoneLog.ToolTipText = ""
37460   Exit Sub
37470 End If

37480 Select Case grdPhoneLog.TextMatrix(0, grdPhoneLog.MouseCol)
        Case "Hae": s = "Haematology"
37490   Case "Bio": s = "Biochemistry"
37500   Case "Coa": s = "Coagulation"
37510   Case "Imm": s = "Immunology"
37520   Case "Gas": s = "Blood Gas"
37530   Case "Ext": s = "External"
37540   Case "Mic": s = "Microbiology"
37550   Case "End": s = "Endocrinology"
37560   Case Else: s = ""
37570 End Select

37580 grdPhoneLog.ToolTipText = s

End Sub



Public Property Let PhoneOrFAX(ByVal strNewValue As String)

37590 pPhoneOrFAX = strNewValue

End Property

