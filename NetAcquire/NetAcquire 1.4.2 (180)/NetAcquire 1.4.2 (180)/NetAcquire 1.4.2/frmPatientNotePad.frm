VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPatientNotePad 
   Caption         =   "NetAcquire - Patient NotePad"
   ClientHeight    =   9540
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   13875
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame framDemographics 
      Caption         =   "Demographics"
      Height          =   1332
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   11712
      Begin VB.Label lblForeNameD 
         AutoSize        =   -1  'True
         Caption         =   "ForeName:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   4020
         TabIndex        =   26
         Top             =   600
         Width           =   816
      End
      Begin VB.Label lblDOBD 
         AutoSize        =   -1  'True
         Caption         =   "DOB:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   6600
         TabIndex        =   25
         Top             =   600
         Width           =   384
      End
      Begin VB.Label lblAgeD 
         AutoSize        =   -1  'True
         Caption         =   "Age :"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   6600
         TabIndex        =   24
         Top             =   900
         Width           =   372
      End
      Begin VB.Label lblChartNoD 
         AutoSize        =   -1  'True
         Caption         =   "Chart Number:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   1500
         TabIndex        =   23
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lblAddressD 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   4020
         TabIndex        =   22
         Top             =   900
         Width           =   684
      End
      Begin VB.Label lblDemoDateD 
         AutoSize        =   -1  'True
         Caption         =   "Demographics Date:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   9720
         TabIndex        =   21
         Top             =   600
         Width           =   1488
      End
      Begin VB.Label lblSexD 
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   6600
         TabIndex        =   20
         Top             =   300
         Width           =   312
      End
      Begin VB.Label lblSurNameD 
         AutoSize        =   -1  'True
         Caption         =   "SurName:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   4020
         TabIndex        =   19
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblSampleIDD 
         AutoSize        =   -1  'True
         Caption         =   "SampleID:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   1500
         TabIndex        =   18
         Top             =   300
         Width           =   756
      End
      Begin VB.Label lblSampleDateD 
         AutoSize        =   -1  'True
         Caption         =   "Sample Date:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   9720
         TabIndex        =   17
         Top             =   300
         Width           =   984
      End
      Begin VB.Label lblSampleDate 
         AutoSize        =   -1  'True
         Caption         =   "Sample Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   8556
         TabIndex        =   16
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label lblSample 
         AutoSize        =   -1  'True
         Caption         =   "SampleID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   552
         TabIndex        =   15
         Top             =   300
         Width           =   876
      End
      Begin VB.Label lblSurName 
         AutoSize        =   -1  'True
         Caption         =   "SurName:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   3096
         TabIndex        =   14
         Top             =   300
         Width           =   828
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   6192
         TabIndex        =   13
         Top             =   300
         Width           =   372
      End
      Begin VB.Label lblDemoDate 
         AutoSize        =   -1  'True
         Caption         =   "Demographics Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   7980
         TabIndex        =   12
         Top             =   600
         Width           =   1716
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   3120
         TabIndex        =   11
         Top             =   900
         Width           =   804
      End
      Begin VB.Label lblChartNo 
         AutoSize        =   -1  'True
         Caption         =   "Chart Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1188
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         Caption         =   "Age :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   6120
         TabIndex        =   9
         Top             =   900
         Width           =   444
      End
      Begin VB.Label lblDOB 
         AutoSize        =   -1  'True
         Caption         =   "DOB:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   6120
         TabIndex        =   8
         Top             =   600
         Width           =   444
      End
      Begin VB.Label lblForeName 
         AutoSize        =   -1  'True
         Caption         =   "ForeName:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   2988
         TabIndex        =   7
         Top             =   600
         Width           =   936
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   1000
      Left            =   12480
      Picture         =   "frmPatientNotePad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   1740
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   420
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmPatientNotePad.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grid(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAddComments(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtComments(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdDelete(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdEdit(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Biochemistry"
      TabPicture(1)   =   "frmPatientNotePad.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdEdit(1)"
      Tab(1).Control(1)=   "cmdDelete(1)"
      Tab(1).Control(2)=   "cmdAddComments(1)"
      Tab(1).Control(3)=   "txtComments(1)"
      Tab(1).Control(4)=   "grid(1)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Haematology"
      TabPicture(2)   =   "frmPatientNotePad.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdEdit(2)"
      Tab(2).Control(1)=   "cmdDelete(2)"
      Tab(2).Control(2)=   "cmdAddComments(2)"
      Tab(2).Control(3)=   "txtComments(2)"
      Tab(2).Control(4)=   "grid(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Microbiology"
      TabPicture(3)   =   "frmPatientNotePad.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grid(3)"
      Tab(3).Control(1)=   "txtComments(3)"
      Tab(3).Control(2)=   "cmdAddComments(3)"
      Tab(3).Control(3)=   "cmdEdit(3)"
      Tab(3).Control(4)=   "cmdDelete(3)"
      Tab(3).ControlCount=   5
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   612
         Index           =   3
         Left            =   -62700
         TabIndex        =   43
         Top             =   1560
         Width           =   1152
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   612
         Index           =   3
         Left            =   -62700
         TabIndex        =   42
         Top             =   900
         Width           =   1152
      End
      Begin VB.CommandButton cmdAddComments 
         Caption         =   "Add Comment"
         Height          =   1000
         Index           =   3
         Left            =   -63840
         Picture         =   "frmPatientNotePad.frx":0D3A
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   6540
         Width           =   1100
      End
      Begin VB.TextBox txtComments 
         Height          =   1332
         Index           =   3
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   6208
         Width           =   10935
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   612
         Index           =   2
         Left            =   -62700
         TabIndex        =   38
         Top             =   900
         Width           =   1152
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   612
         Index           =   2
         Left            =   -62700
         TabIndex        =   37
         Top             =   1560
         Width           =   1152
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   612
         Index           =   1
         Left            =   -62700
         TabIndex        =   36
         Top             =   900
         Width           =   1152
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   612
         Index           =   1
         Left            =   -62700
         TabIndex        =   35
         Top             =   1560
         Width           =   1152
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   612
         Index           =   0
         Left            =   12300
         TabIndex        =   34
         Top             =   900
         Width           =   1152
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   612
         Index           =   0
         Left            =   12300
         TabIndex        =   33
         Top             =   1560
         Width           =   1152
      End
      Begin VB.CommandButton cmdAddComments 
         Caption         =   "Add Comment"
         Height          =   1000
         Index           =   2
         Left            =   -63840
         Picture         =   "frmPatientNotePad.frx":1604
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6540
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddComments 
         Caption         =   "Add Comment"
         Height          =   1000
         Index           =   1
         Left            =   -63840
         Picture         =   "frmPatientNotePad.frx":1ECE
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6540
         Width           =   1100
      End
      Begin VB.TextBox txtComments 
         Height          =   1332
         Index           =   2
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   6208
         Width           =   10935
      End
      Begin VB.TextBox txtComments 
         Height          =   1332
         Index           =   1
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   6208
         Width           =   10935
      End
      Begin VB.TextBox txtComments 
         Height          =   1332
         Index           =   0
         Left            =   120
         MaxLength       =   3999
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   6208
         Width           =   10935
      End
      Begin VB.CommandButton cmdAddComments 
         Caption         =   "Add Comment"
         Height          =   1000
         Index           =   0
         Left            =   11160
         Picture         =   "frmPatientNotePad.frx":2798
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6540
         Width           =   1100
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   5535
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   10
         WordWrap        =   -1  'True
         FormatString    =   $"frmPatientNotePad.frx":3062
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   5535
         Index           =   1
         Left            =   -74880
         TabIndex        =   27
         Top             =   600
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   10
         WordWrap        =   -1  'True
         FormatString    =   $"frmPatientNotePad.frx":3141
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   5535
         Index           =   2
         Left            =   -74880
         TabIndex        =   28
         Top             =   600
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   10
         WordWrap        =   -1  'True
         FormatString    =   $"frmPatientNotePad.frx":3220
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   5535
         Index           =   3
         Left            =   -74880
         TabIndex        =   39
         Top             =   600
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   10
         WordWrap        =   -1  'True
         FormatString    =   $"frmPatientNotePad.frx":32FF
      End
   End
   Begin VB.Label lbl1 
      Caption         =   "Previous Comments:"
      Height          =   252
      Left            =   180
      TabIndex        =   6
      Top             =   1500
      Width           =   1692
   End
End
Attribute VB_Name = "frmPatientNotePad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SampleID As String
Public Caller As String
Public GLabNo As String

Private Sub cmdAddComments_Click(Index As Integer)
          Dim sql As String

31750     On Error GoTo cmdAddComments_Click_Error

31760     If Trim(txtComments(Index)) = "" Then Exit Sub
31770     Comment = txtComments(Index).Text

31780     sql = "INSERT INTO PatientNotePad " & _
              "(SampleID, DateTimeofRecord, Comment, UserName, Descipline, LabNo)" & _
              "VALUES ( '" & lblSampleIDD & "', FORMAT(GETDATE(), 'dd/MMM/yyyy HH:mm:ss'),'" & WordWrap(Replace(txtComments(Index).Text, vbTab, " "), 185) & "','" & UserName & "','" & SSTab.TabCaption(SSTab.Tab) & "','" & Val(GLabNo) & "')"
31790     Cnxn(0).Execute sql
31800     grid(Index).AddItem Format$(Now, "dd/mmm/yyyy hh:mm:ss") & vbTab & WordWrap(Replace(txtComments(Index).Text, vbTab, " "), 185) & vbTab & UserName & vbTab & lblSampleIDD.Caption, 1
31810     grid(Index).RowHeight(1) = TextHeight(WordWrap(txtComments(Index).Text, 185))
31820     txtComments(Index) = ""
          'cmdAddComments(Index).Enabled = False

31830     Exit Sub

cmdAddComments_Click_Error:

          Dim strES As String
          Dim intEL As Integer

31840     intEL = Erl
31850     strES = Err.Description
31860     LogError "frmPatientNotePad", "cmdAddComments_Click", intEL, strES, sql
          
End Sub

Private Sub cmdClose_Click()
31870 On Error GoTo cmdClose_Click_Error

31880 Unload Me

31890 Exit Sub

cmdClose_Click_Error:

       Dim strES As String
       Dim intEL As Integer

31900  intEL = Erl
31910  strES = Err.Description
31920  LogError "frmPatientNotePad", "cmdClose_Click", intEL, strES
          
End Sub

Private Function WordWrap(strFullText As String, intLength As Integer) As String

      Dim intLen As Integer, intCr As Integer, intSpace As Integer
      Dim strText As String, strNextLine As String
      Dim blnDoneOnce As Boolean

31930 On Error GoTo WordWrap_Error

31940 intLength = intLength + 1
31950 strFullText = Trim$(strFullText)

31960 Do
31970     intLen = Len(strNextLine)
31980     intSpace = InStr(strFullText, " ")
31990     intCr = InStr(strFullText, vbCr)

32000     If intCr Then
32010         If intLen + intCr <= intLength Then
32020             strText = strText & strNextLine & Left$(strFullText, intCr)
32030             strNextLine = ""
32040             strFullText = Mid$(strFullText, intCr + 1)
32050             GoTo LoopHere
32060         End If
32070     End If

32080     If intSpace Then
32090         If intLen + intSpace <= intLength Then
32100             blnDoneOnce = True
32110             strNextLine = strNextLine & Left$(strFullText, intSpace)
32120             strFullText = Mid$(strFullText, intSpace + 1)
32130         ElseIf intSpace > intLength Then
32140             strText = strText & vbCrLf & Left$(strFullText, intLength)
32150             strFullText = Mid$(strFullText, intLength + 1)
32160         Else
32170             strText = strText & strNextLine & vbCrLf
32180             strNextLine = ""
32190         End If
32200     Else
32210         If intLen Then
32220             If intLen + Len(strFullText) > intLength Then
32230                 strText = strText & strNextLine & vbCrLf & strFullText & vbCrLf
32240             Else
32250                 strText = strText & strNextLine & strFullText & vbCrLf
32260             End If
32270         Else
32280             strText = strText & strFullText & vbCrLf
32290         End If
32300         Exit Do
32310     End If

LoopHere:
32320 Loop

32330 WordWrap = strText

32340 Exit Function

WordWrap_Error:

       Dim strES As String
       Dim intEL As Integer

32350  intEL = Erl
32360  strES = Err.Description
32370  LogError "frmPatientNotePad", "WordWrap", intEL, strES
          

End Function

Private Sub LoadComments()



32380 On Error GoTo LoadComments_Error

32390 If Val(GLabNo) = 0 Then
32400     If Caller = "Microbiology" Then
32410         LoadDescipline Val(SampleID), SSTab.TabCaption(SSTab.Tab), 3
32420     Else
32430         LoadDescipline Val(SampleID), SSTab.TabCaption(0), 0
32440         LoadDescipline Val(SampleID), SSTab.TabCaption(1), 1
32450         LoadDescipline Val(SampleID), SSTab.TabCaption(2), 2
32460     End If
32470 Else
32480     LoadDescipline Val(SampleID), SSTab.TabCaption(0), 0, GLabNo
32490     LoadDescipline Val(SampleID), SSTab.TabCaption(1), 1, GLabNo
32500     LoadDescipline Val(SampleID), SSTab.TabCaption(2), 2, GLabNo
32510     LoadDescipline Val(SampleID), SSTab.TabCaption(3), 3, GLabNo
32520 End If

32530 Exit Sub

LoadComments_Error:

       Dim strES As String
       Dim intEL As Integer

32540  intEL = Erl
32550  strES = Err.Description
32560  LogError "frmPatientNotePad", "LoadComments", intEL, strES
          
End Sub

Private Sub LoadDemo(SampleID As String)
      Dim sql As String
      Dim tb As Recordset

32570 On Error GoTo LoadDemo_Error

32580 sql = "Select * from Demographics where " & _
            "SampleID = '" & SampleID & "'"

32590 Set tb = New Recordset
32600 RecOpenClient 0, tb, sql
32610 With tb
32620     If Not tb.EOF Then
32630         lblSampleIDD = !SampleID
32640         lblChartNoD = !Chart & ""
32650         lblSurNameD = SurName(!PatName & "")
32660         lblForeNameD = ForeName(!PatName & "")
32670         lblAddressD = !Addr0 & " " & !Addr1 & ""
32680         lblSexD = !Sex & ""
32690         lblDOBD = !DoB
32700         lblAgeD = !Age
32710         lblSampleDateD = !SampleDate
32720         lblDemoDateD = !DateTimeDemographics
32730         GLabNo = !LabNo & ""

32740     End If
32750 End With

32760 Exit Sub

LoadDemo_Error:

      Dim strES As String
      Dim intEL As Integer

32770 intEL = Erl
32780 strES = Err.Description
32790 LogError "frmPatientNotePad", "LoadDemo", intEL, strES, sql

End Sub
Private Sub LoadDescipline(SampleID As String, Descipline As String, Index As Integer, Optional LabNo As String = "")
          Dim sql As String
          Dim tb As Recordset
          Dim LNo As Long

          'cmdEdit(Index).Enabled = False
          'cmdDelete(Index).Enabled = False
          'cmdAddComments(Index).Enabled = False
32800     On Error GoTo LoadDescipline_Error




32810     grid(Index).Clear
32820     grid(Index).Rows = 1

32830     grid(Index).SelectionMode = flexSelectionByRow


32840     If LabNo = "" Then
32850         sql = "Select * from PatientNotePad where " & _
                    "SampleID = '" & SampleID & "' and " & _
                    "Descipline = '" & Descipline & "' ORDER BY DateTimeOfRecord DESC"
32860     Else
32870         sql = "Select * from PatientNotePad where " & _
                    "labNo = '" & LabNo & "' and " & _
                    "Descipline = '" & Descipline & "' ORDER BY DateTimeOfRecord DESC"
32880     End If
32890     Set tb = New Recordset
32900     RecOpenClient 0, tb, sql
32910     With tb
32920         If tb.EOF Then
32930         Else
32940             Do Until .EOF
32950                 grid(Index).AddItem Format$(!DateTimeOfRecord, "dd/MMM/yyyy HH:mm:ss") & vbTab & WordWrap(Replace(!Comment & "", vbTab, " "), 90) & vbTab & !UserName & vbTab & !SampleID
32960                 grid(Index).RowHeight(grid(Index).Rows - 1) = TextHeight(WordWrap(!Comment & "", 90))
32970                 .MoveNext
32980             Loop
32990         End If
33000     End With

33010     Exit Sub

LoadDescipline_Error:

          Dim strES As String
          Dim intEL As Integer

33020     intEL = Erl
33030     strES = Err.Description
33040     LogError "frmPatientNotePad", "LoadDescipline", intEL, strES, sql


End Sub

Private Sub cmdDelete_Click(Index As Integer)
      Dim sql As String

33050 On Error GoTo cmdDelete_Click_Error

33060 If grid(Index).row = 0 Then Exit Sub

33070 If iMsg("Do you want to delete selected comment", vbYesNo) = vbNo Then Exit Sub
          

33080 sql = "delete from PatientNotePad where Sampleid = '" & SampleID & "' and DateTimeofRecord = '" & grid(Index).TextMatrix(grid(Index).row, 0) & "'"
33090 Cnxn(0).Execute sql

33100 If grid(Index).Rows = 2 Then
33110     grid(Index).Clear
33120     grid(Index).Rows = 1
33130 Else
33140     grid(Index).RemoveItem (grid(Index).row)
33150 End If
33160 grid(Index).row = 0
      'cmdDelete(Index).Enabled = False
      'cmdEdit(Index).Enabled = False

33170 Exit Sub

cmdDelete_Click_Error:

       Dim strES As String
       Dim intEL As Integer

33180  intEL = Erl
33190  strES = Err.Description
33200  LogError "frmPatientNotePad", "cmdDelete_Click", intEL, strES
          

End Sub

Private Sub cmdEdit_Click(Index As Integer)
33210 On Error GoTo cmdEdit_Click_Error

33220 txtComments(Index) = grid(Index).TextMatrix(grid(Index).row, 1)
33230 If grid(Index).Rows = 2 Then
33240     grid(Index).Clear
33250     grid(Index).Rows = 1
33260 Else
33270     grid(Index).RemoveItem (grid(Index).row)
33280 End If
33290 grid(Index).row = 0
      'cmdEdit(Index).Enabled = False
      'cmdDelete(Index).Enabled = False
33300 txtComments(Index).SetFocus

33310 Exit Sub

cmdEdit_Click_Error:

       Dim strES As String
       Dim intEL As Integer

33320  intEL = Erl
33330  strES = Err.Description
33340  LogError "frmPatientNotePad", "cmdEdit_Click", intEL, strES
          
End Sub

Private Sub Form_Load()
          Dim sql As String
33350     On Error GoTo Form_Load_Error


33360     For i = 0 To 3
33370         grid(i).Clear
33380         grid(i).Rows = 1
33390     Next i

33400     CheckPatientNotepadInDb

33410     If Caller = "Microbiology" Then
33420         SampleID = Val(SampleID) ' + sysOptMicroOffset(0)
33430         SSTab.Tab = 3
              'Disable other discipline changes
33440         LockForm True
33450         SSTab.TabVisible(0) = False
33460         SSTab.TabVisible(1) = False
33470         SSTab.TabVisible(2) = False
              
33480     Else
33490         SSTab.Tab = 0
33500         LockForm False
33510         SSTab.TabVisible(3) = False
33520     End If

33530     LoadDemo SampleID
33540     LoadComments



33550     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

33560     intEL = Erl
33570     strES = Err.Description
33580     LogError "frmPatientNotePad", "Form_Load", intEL, strES, sql

End Sub



Private Sub LockForm(ByVal EnableMicro As Boolean)

      Dim i As Integer

33590 On Error GoTo LockControls_Error


33600 For i = 0 To 2
33610     cmdAddComments(i).Enabled = Not EnableMicro
33620     cmdEdit(i).Enabled = Not EnableMicro
33630     cmdDelete(i).Enabled = Not EnableMicro
33640 Next i
33650 cmdAddComments(3).Enabled = EnableMicro
33660 cmdEdit(3).Enabled = EnableMicro
33670 cmdDelete(3).Enabled = EnableMicro

33680 Exit Sub

LockControls_Error:

       Dim strES As String
       Dim intEL As Integer

33690  intEL = Erl
33700  strES = Err.Description
33710  LogError "frmPatientNotePad", "LockForm", intEL, strES

End Sub

Private Sub grid_Click(Index As Integer)
      Dim CurRow As Integer
33720 On Error GoTo grid_Click_Error

33730 CurRow = grid(Index).MouseRow
      'If CurRow > 0 Then
      '    cmdEdit(Index).Enabled = True
      '    cmdDelete(Index).Enabled = True
      'End If

33740 Exit Sub

grid_Click_Error:

       Dim strES As String
       Dim intEL As Integer

33750  intEL = Erl
33760  strES = Err.Description
33770  LogError "frmPatientNotePad", "grid_Click", intEL, strES
          
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
33780 On Error GoTo SSTab_Click_Error

33790 LoadComments

33800 Exit Sub

SSTab_Click_Error:

       Dim strES As String
       Dim intEL As Integer

33810  intEL = Erl
33820  strES = Err.Description
33830  LogError "frmPatientNotePad", "SSTab_Click", intEL, strES
          
End Sub

'Private Sub txtComments_Change(Index As Integer)
'cmdAddComments(Index).Enabled = Len(txtComments(Index))
'End Sub
