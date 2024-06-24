VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAudit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   13590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   915
      Left            =   11640
      ScaleHeight     =   855
      ScaleWidth      =   1725
      TabIndex        =   6
      Top             =   6540
      Width           =   1785
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Deletion"
         Height          =   195
         Left            =   570
         TabIndex        =   12
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Changes Made"
         Height          =   195
         Left            =   570
         TabIndex        =   11
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Red"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   330
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Changes"
         Height          =   195
         Left            =   570
         TabIndex        =   8
         Top             =   60
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Green"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   435
      End
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   5415
      Left            =   11640
      TabIndex        =   5
      Top             =   1050
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   9551
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtSampleID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11647
      TabIndex        =   3
      Top             =   600
      Width           =   1770
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   8475
      Left            =   60
      TabIndex        =   2
      Top             =   270
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   14949
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAudit.frx":0000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1100
      Left            =   11917
      Picture         =   "frmAudit.frx":008B
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   7650
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView ListViewCodes 
      Height          =   5415
      Left            =   10770
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   9551
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   12150
      TabIndex        =   4
      Top             =   390
      Width           =   765
   End
End
Attribute VB_Name = "frmAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pTableName As String
Private pTableNameAudit As String

Private Dept As String
Private Sub LoadAudit()

          Dim sql As String
          Dim tb As Recordset
          Dim tbArc As Recordset
          Dim n As Integer
          Dim CurrentName() As String
          Dim Current() As String
          Dim NameDisplayed As Boolean

57780     On Error GoTo LoadAudit_Error

57790     rtb.Text = ""
57800     rtb.SelFontSize = 12

57810     If Trim$(txtSampleID) = "" Then Exit Sub

57820     rtb.SelFontSize = 16
57830     rtb.SelColor = vbBlack
57840     rtb.SelUnderline = True
57850     rtb.SelBold = True
57860     rtb.SelText = "Audit Trail for "
57870     rtb.SelColor = vbRed
57880     rtb.SelText = IIf(InStr(1, pTableName, "Coag"), CoagNameFor(ListView.SelectedItem.Text), ListView.SelectedItem.Text) & vbCrLf & vbCrLf
57890     rtb.SelUnderline = False

57900     rtb.SelFontSize = 12

57910     sql = "SELECT * FROM " & pTableName & " WHERE " & _
              "SampleID = '" & txtSampleID & "' " & _
              "AND Code = '" & ListViewCodes.ListItems(ListView.SelectedItem.Index) & "' "
          '      "  (SELECT Distinct Code FROM " & Dept & "TestDefinitions WHERE " & _
          '      "   ShortName = '" & ListView.SelectedItem.Text & "')"
57920     If InStr(1, pTableName, "Coag") > 0 Then
57930         sql = Replace(sql, "ShortName", "Code")
57940     End If
57950     Set tb = New Recordset
57960     RecOpenServer 0, tb, sql
57970     If tb.EOF Then
57980         rtb.SelText = "No Current Record found." & vbCrLf
57990     End If

58000     ReDim Current(0 To tb.Fields.Count - 1)
58010     ReDim CurrentName(0 To tb.Fields.Count - 1)
58020     For n = 0 To tb.Fields.Count - 1
58030         If Not tb.EOF Then
58040             Current(n) = tb.Fields(n).Value & ""
58050         Else
58060             Current(n) = ""
58070         End If
58080         CurrentName(n) = tb.Fields(n).Name

58090         sql = "SELECT ArchivedBy, ArchiveDateTime, [" & CurrentName(n) & "] FROM " & pTableNameAudit & " WHERE " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "AND Code = '" & ListViewCodes.ListItems(ListView.SelectedItem.Index) & "' " & _
                  "ORDER BY ArchiveDateTime DESC"
              '"  (SELECT DISTINCT Code FROM " & Dept & "TestDefinitions WHERE " & _
              '"   ShortName = '" & ListView.SelectedItem.Text & "') " &
58100         If InStr(1, pTableNameAudit, "Coag") > 0 Then
58110             sql = Replace(sql, "ShortName", "Code")
58120         End If
58130         Set tbArc = New Recordset
58140         RecOpenServer 0, tbArc, sql
58150         If tbArc.EOF Then
58160             rtb.SelText = "No Changes Made."
58170             Exit For
58180         Else
58190             NameDisplayed = False
58200             Do While Not tbArc.EOF
58210                 If Trim$(Current(n)) <> Trim$(tbArc.Fields(CurrentName(n)) & "") Then
58220                     If Not NameDisplayed Then
58230                         rtb.SelBold = True
58240                         rtb.SelColor = vbBlue
58250                         rtb.SelFontSize = 12
58260                         rtb.SelText = CurrentName(n) & vbCrLf
58270                         NameDisplayed = True
58280                     End If
58290                     rtb.SelFontSize = 12
58300                     rtb.SelText = tbArc!ArchiveDateTime & " "
58310                     rtb.SelColor = vbRed
58320                     rtb.SelText = tbArc!ArchivedBy & ""
58330                     rtb.SelColor = vbBlack
58340                     rtb.SelText = " Changed "
58350                     rtb.SelColor = vbGreen
58360                     rtb.SelBold = True
58370                     If Trim$(tbArc.Fields(CurrentName(n)) & "") = "" Then
58380                         rtb.SelText = "<Blank> "
58390                     Else
58400                         rtb.SelText = Trim$(tbArc.Fields(CurrentName(n)))
58410                     End If
58420                     rtb.SelColor = vbBlack
58430                     rtb.SelBold = False
58440                     rtb.SelText = " to "
58450                     rtb.SelBold = True
58460                     If Trim$(Current(n)) = "" Then
58470                         rtb.SelText = "<Blank>" & vbCrLf
58480                     Else
58490                         rtb.SelText = Trim$(Current(n)) & vbCrLf
58500                     End If
58510                     Current(n) = Trim$(tbArc.Fields(CurrentName(n)) & "")
58520                 End If
58530                 tbArc.MoveNext
58540             Loop
58550         End If
58560         If NameDisplayed Then
58570             rtb.SelText = vbCrLf
58580         End If
58590     Next

58600     Exit Sub

LoadAudit_Error:

          Dim strES As String
          Dim intEL As Integer

58610     intEL = Erl
58620     strES = Err.Description
58630     LogError "frmAudit", "LoadAudit", intEL, strES, sql

End Sub

Private Sub SelectDisplay()

          Dim sql As String
          Dim tb As Recordset
          Dim clmX As ColumnHeader
          Dim itmX As MSComctlLib.ListItem
          Dim itmC As MSComctlLib.ListItem

58640     On Error GoTo SelectDisplay_Error

58650     rtb.TextRTF = ""

58660     ListView.ListItems.Clear
58670     Set clmX = ListView.ColumnHeaders.Add()
58680     clmX.Text = "Parameter"

58690     txtSampleID = Val(txtSampleID)
58700     If Val(txtSampleID) = 0 Then Exit Sub

58710     Select Case UCase$(pTableName)
              Case "BIORESULTS": Dept = "Bio"
58720         Case "ENDRESULTS": Dept = "End"
58730         Case "COAGRESULTS": Dept = "Coag"
58740     End Select

58750     If Dept <> "" Then

58760         sql = "SELECT DISTINCT ShortName, Code Cod, '65280' Colour FROM " & Dept & "TestDefinitions WHERE " & _
                  "  Code IN (  SELECT DISTINCT Code FROM " & Dept & "Results WHERE " & _
                  "             SampleID = '" & txtSampleID & "' ) " & _
                  "  AND Code NOT IN ( SELECT DISTINCT Code FROM " & Dept & "ResultsAudit WHERE " & _
                  "           SampleID = '" & txtSampleID & "' ) " & _
                  "UNION " & _
                  "SELECT DISTINCT ShortName, Code Cod, '255' Colour FROM " & Dept & "TestDefinitions WHERE " & _
                  "  Code IN ( SELECT DISTINCT Code FROM " & Dept & "ResultsAudit WHERE " & _
                  "            SampleID = '" & txtSampleID & "' ) " & _
                  "  AND Code NOT IN (  SELECT DISTINCT Code FROM " & Dept & "Results WHERE " & _
                  "             SampleID = '" & txtSampleID & "' ) " & _
                  "UNION " & _
                  "SELECT DISTINCT ShortName, Code Cod, '16711680' Colour FROM " & Dept & "TestDefinitions WHERE " & _
                  "  Code IN ( SELECT DISTINCT Code FROM " & Dept & "ResultsAudit WHERE " & _
                  "            SampleID = '" & txtSampleID & "' ) " & _
                  "  AND Code IN (  SELECT DISTINCT Code FROM " & Dept & "Results WHERE " & _
                  "             SampleID = '" & txtSampleID & "' ) " & _
                  "GROUP BY ShortName, Code ORDER BY ShortName"
58770         If Dept = "Coag" Then
58780             sql = Replace(sql, "ShortName", "Code")
58790         End If
58800         Set tb = New Recordset
58810         RecOpenServer 0, tb, sql
58820         If Not tb.EOF Then
58830             Do While Not tb.EOF

58840                 Set itmX = ListView.ListItems.Add()
58850                 Set itmC = ListViewCodes.ListItems.Add()
58860                 If Dept = "Coag" Then
58870                     itmX.Text = CoagNameFor(tb!Code) & ""
58880                 Else
58890                     itmX.Text = tb!ShortName & ""
58900                 End If
58910                 itmC.Text = tb!Cod & ""
58920                 itmX.ForeColor = tb!Colour

58930                 tb.MoveNext
58940             Loop
58950             If ListView.ListItems.Count > 0 Then
58960                 ListView.ListItems(1).Selected = True
58970                 LoadAudit
58980             End If
58990         Else
59000             rtb.SelText = "No Record."
59010         End If

59020     End If

59030     Exit Sub

SelectDisplay_Error:

          Dim strES As String
          Dim intEL As Integer

59040     intEL = Erl
59050     strES = Err.Description
59060     LogError "frmAudit", "SelectDisplay", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

59070     Unload Me

End Sub


Private Sub Form_Activate()

59080     SelectDisplay

End Sub
Public Property Let TableName(ByVal sNewValue As String)

59090     pTableName = sNewValue
59100     pTableNameAudit = sNewValue & "Audit"

End Property
Public Property Let SampleID(ByVal sNewValue As String)

59110     txtSampleID = sNewValue

End Property

Private Sub ListView_Click()

59120     LoadAudit

End Sub

Private Sub ListView_KeyPress(KeyAscii As Integer)

59130     KeyAscii = 0

End Sub


Private Sub txtsampleid_LostFocus()

59140     rtb.TextRTF = ""

59150     SelectDisplay

End Sub

