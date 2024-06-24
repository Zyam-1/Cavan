VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAutoGenerateCommentsMicro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   14055
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   1000
         Left            =   12720
         Picture         =   "frmAutoGenerateCommentsMicro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "bAdd"
         Top             =   3720
         Width           =   1100
      End
      Begin VB.CommandButton bCancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   1000
         Left            =   12720
         Picture         =   "frmAutoGenerateCommentsMicro.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   5880
         Width           =   1100
      End
      Begin VB.CommandButton cmdSaveOrders 
         Caption         =   "&Save"
         Height          =   1000
         Left            =   10440
         Picture         =   "frmAutoGenerateCommentsMicro.frx":16EC
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   420
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Reset"
         Height          =   1000
         Left            =   10440
         Picture         =   "frmAutoGenerateCommentsMicro.frx":306E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1500
         Width           =   1100
      End
      Begin VB.TextBox txtCounter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CheckBox chkPhoneAlert 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone Alert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5640
         TabIndex        =   21
         Top             =   2040
         Width           =   1965
      End
      Begin VB.TextBox txtAgeToDays 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7320
         TabIndex        =   8
         Top             =   480
         Width           =   765
      End
      Begin VB.TextBox txtAgeFromDays 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6120
         TabIndex        =   6
         Top             =   480
         Width           =   705
      End
      Begin VB.TextBox txtComment 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         MaxLength       =   400
         TabIndex        =   22
         Top             =   2160
         Width           =   4395
      End
      Begin VB.ComboBox cmbWard 
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Text            =   "cmbWard"
         Top             =   1560
         Width           =   3060
      End
      Begin VB.ComboBox cmbOrgName 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cmbOrgName"
         Top             =   480
         Width           =   3045
      End
      Begin VB.ComboBox cmbSite 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   12
         Text            =   "cmbSite"
         Top             =   1020
         Width           =   3045
      End
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   315
         Left            =   6180
         TabIndex        =   15
         Top             =   1380
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   220200961
         CurrentDate     =   40093
      End
      Begin MSComCtl2.DTPicker dtEnd 
         Height          =   315
         Left            =   8280
         TabIndex        =   17
         Top             =   1380
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   220200961
         CurrentDate     =   40093
      End
      Begin MSComCtl2.DTPicker dtSchDate 
         Height          =   315
         Left            =   7620
         TabIndex        =   24
         Top             =   2400
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         Format          =   220200961
         CurrentDate     =   40093
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   3315
         Left            =   180
         TabIndex        =   25
         Top             =   3660
         Width           =   12420
         _ExtentX        =   21908
         _ExtentY        =   5847
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
         HighLight       =   2
         GridLines       =   2
         SelectionMode   =   1
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator|<Code"
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Scheduled Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   23
         Top             =   2460
         Width           =   1380
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6900
         TabIndex        =   7
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   5
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Age ( Day(s) )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   1
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Comment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(Maximum 400 characters)"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1140
         TabIndex        =   20
         Top             =   1920
         Width           =   1860
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "This rule is active between"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   11
         Top             =   1020
         Width           =   2310
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7920
         TabIndex        =   16
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5700
         TabIndex        =   14
         Top             =   1380
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Org Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Site"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   825
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmAutoGenerateCommentsMicro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Procedure : FillSites
' Author    : Masood
' Date      : 08/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub FillSites()

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String


61720     On Error GoTo FillSites_Error


61730     cmbSite.Clear
61740     sql = "SELECT Text, UPPER(ListType) ListType FROM Lists " & _
              "WHERE " & _
              "ListType IN ( 'SI') " & _
              "AND InUse = 1 " & _
              "ORDER BY ListOrder"
61750     Set tb = New Recordset
61760     RecOpenClient 0, tb, sql
61770     Do While Not tb.EOF
61780         cmbSite.AddItem tb!Text & ""
61790         tb.MoveNext
61800     Loop

61810     cmbSite.AddItem ("Any")

61820     Exit Sub


FillSites_Error:

          Dim strES As String
          Dim intEL As Integer

61830     intEL = Erl
61840     strES = Err.Description
61850     LogError "frmAutoGenerateCommentsMicro", "FillSites", intEL, strES, sql

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FillOrgNames
' Author    : Masood
' Date      : 08/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub FillOrgNames()

          Dim tb As Recordset
          Dim sql As String


61860     On Error GoTo FillOrgNames_Error


61870     cmbOrgName.Clear

61880     sql = "Select * from Organisms  " & _
              " order by ListOrder"
          '  "GroupName = '" & cmbOrgGroup(Index).Text & "' "
61890     Set tb = New Recordset
61900     RecOpenClient 0, tb, sql
61910     Do While Not tb.EOF
61920         cmbOrgName.AddItem tb!Name & ""
61930         tb.MoveNext
61940     Loop

61950     cmbOrgName.AddItem ("Any")

61960     Exit Sub


FillOrgNames_Error:

          Dim strES As String
          Dim intEL As Integer

61970     intEL = Erl
61980     strES = Err.Description
61990     LogError "frmAutoGenerateCommentsMicro", "FillOrgNames", intEL, strES, sql

End Sub


Private Sub bcancel_Click()
62000     Unload Me
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbOrgName_KeyPress
' Author    : Masood
' Date      : 14/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbOrgName_KeyPress(KeyAscii As Integer)
62010     On Error GoTo cmbOrgName_KeyPress_Error


62020     KeyAscii = AutoComplete(cmbOrgName, KeyAscii, False)


62030     Exit Sub


cmbOrgName_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

62040     intEL = Erl
62050     strES = Err.Description
62060     LogError "frmAutoGenerateCommentsMicro", "cmbOrgName_KeyPress", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbSite_KeyPress
' Author    : Masood
' Date      : 14/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbSite_KeyPress(KeyAscii As Integer)
62070     On Error GoTo cmbSite_KeyPress_Error


62080     KeyAscii = AutoComplete(cmbSite, KeyAscii, False)


62090     Exit Sub


cmbSite_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

62100     intEL = Erl
62110     strES = Err.Description
62120     LogError "frmAutoGenerateCommentsMicro", "cmbSite_KeyPress", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbWard_KeyPress
' Author    : Masood
' Date      : 14/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbWard_KeyPress(KeyAscii As Integer)
62130     On Error GoTo cmbWard_KeyPress_Error


62140     KeyAscii = AutoComplete(cmbWard, KeyAscii, False)


62150     Exit Sub


cmbWard_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

62160     intEL = Erl
62170     strES = Err.Description
62180     LogError "frmAutoGenerateCommentsMicro", "cmbWard_KeyPress", intEL, strES
End Sub

Private Sub cmdAdd_Click()
62190     EnableControls ("N")
62200     cmdSaveOrders.Enabled = True
End Sub

Private Sub cmdClear_Click()
62210     ClearValues
End Sub

Private Sub cmdEdit_Click()
62220     EditRowSelected

End Sub

Private Sub cmdSaveOrders_Click()
62230     SaveComentsMicro
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveComentsMicro
' Author    : Masood
' Date      : 08/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SaveComentsMicro()

          Dim sql As String
          Dim NewQry As String
          Dim UpdateQry As String
62240     On Error GoTo SaveComentsMicro_Error

62250     If cmbOrgName = "" Or cmbOrgName = "Select" Then
62260         iMsg "Please select organism first", vbExclamation
62270         Exit Sub
62280     End If

62290     If txtComment = "" And chkPhoneAlert.Value = 0 Then
62300         iMsg "Please enter comment or select phone alert option", vbExclamation
62310         Exit Sub
62320     End If

62330     NewQry = ""
62340     NewQry = NewQry & " INSERT INTO MicroAutoCommentAlert " & vbNewLine
62350     NewQry = NewQry & " (OrganismName, Site, PatientLocation" & vbNewLine
62360     NewQry = NewQry & " ,DateStart,DateEnd,Comment,PatientAgeFrom,PatientAgeTo,PhoneAlert, PhoneAlertDateTime,ListOrder" & vbNewLine
62370     NewQry = NewQry & " )" & vbNewLine
62380     NewQry = NewQry & " VALUES ( " & vbNewLine
62390     NewQry = NewQry & "@OrganismName,@Site,@PatientLocation,@DateStart,@DateEnd,@Comment,@PatientAgeFrom,@PatientAgeTo" & vbNewLine
62400     NewQry = NewQry & "," & chkPhoneAlert.Value & ",'" & Format(dtSchDate, "yyyy-mm-dd") & "',0" & vbNewLine
62410     NewQry = NewQry & " )" & vbNewLine

62420     UpdateQry = "" & vbNewLine
62430     UpdateQry = UpdateQry & " UPDATE MicroAutoCommentAlert SET " & vbNewLine
62440     UpdateQry = UpdateQry & " OrganismName = @OrganismName" & vbNewLine
62450     UpdateQry = UpdateQry & "  , Site = @Site " & vbNewLine
62460     UpdateQry = UpdateQry & " , PatientLocation = @PatientLocation" & vbNewLine
62470     UpdateQry = UpdateQry & " , PatientAgeFrom = @PatientAgeFrom" & vbNewLine
62480     UpdateQry = UpdateQry & " , PatientAgeTo = @PatientAgeTo" & vbNewLine
62490     UpdateQry = UpdateQry & " , DateStart = @DateStart" & vbNewLine
62500     UpdateQry = UpdateQry & " , DateEnd = @DateEnd" & vbNewLine
62510     UpdateQry = UpdateQry & " , Comment = @Comment" & vbNewLine
62520     UpdateQry = UpdateQry & " , PhoneAlert = " & chkPhoneAlert.Value & vbNewLine
62530     UpdateQry = UpdateQry & " , PhoneAlertDateTime= '" & Format(dtSchDate, "yyyy-mm-dd") & "'" & vbNewLine
62540     UpdateQry = UpdateQry & " WHERE Counter = " & Val(txtCounter) & vbNewLine

62550     sql = "IF EXISTS " & vbNewLine
62560     sql = sql & " ( SELECT * FROM MicroAutoCommentAlert WHERE Counter = " & Val(txtCounter) & " ) " & vbNewLine
62570     sql = sql & " " & UpdateQry & vbNewLine
62580     sql = sql & " ELSE " & vbNewLine
62590     sql = sql & " BEGIN " & vbNewLine
62600     sql = sql & " " & NewQry & vbNewLine
62610     sql = sql & " END " & vbNewLine


62620     If cmbOrgName = "Any" Or cmbOrgName = "" Then
62630         sql = Replace(sql, "@OrganismName", "Null")
62640     Else
62650         sql = Replace(sql, "@OrganismName", "'" & cmbOrgName & "'")
62660     End If


62670     If cmbSite = "Any" Or cmbSite = "" Then
62680         sql = Replace(sql, "@Site", "Null")
62690     Else
62700         sql = Replace(sql, "@Site", "'" & cmbSite & "'")
62710     End If


62720     If cmbWard = "Any" Or cmbWard = "" Then
62730         sql = Replace(sql, "@PatientLocation", "Null")
62740     Else
62750         sql = Replace(sql, "@PatientLocation", "'" & cmbWard & "'")
62760     End If

62770     If IsNull(dtStart) Then
62780         sql = Replace(sql, "@DateStart", "Null")
62790     Else
62800         sql = Replace(sql, "@DateStart", "'" & Format(dtStart, "yyyy-mm-dd") & "'")
62810     End If

62820     If IsNull(dtEnd) Then
62830         sql = Replace(sql, "@DateEnd", "Null")
62840     Else
62850         sql = Replace(sql, "@DateEnd", "'" & Format(dtEnd, "yyyy-mm-dd") & "'")
62860     End If


62870     If Val(txtAgeFromDays) = 0 Then
62880         sql = Replace(sql, "@PatientAgeFrom", "Null")
62890     Else
62900         sql = Replace(sql, "@PatientAgeFrom", "'" & Val(txtAgeFromDays) & "'")
62910     End If

62920     If Val(txtAgeToDays) = 0 Then
62930         sql = Replace(sql, "@PatientAgeTo", "Null")
62940     Else
62950         sql = Replace(sql, "@PatientAgeTo", "'" & Val(txtAgeToDays) & "'")
62960     End If


62970     sql = Replace(sql, "@Comment", "'" & txtComment & "'")




62980     Cnxn(0).Execute sql
          '    iMsg ("Record is saved")
62990     ClearValues
63000     Exit Sub


SaveComentsMicro_Error:

          Dim strES As String
          Dim intEL As Integer

63010     intEL = Erl
63020     strES = Err.Description
63030     LogError "frmAutoGenerateCommentsMicro", "SaveComentsMicro", intEL, strES, sql

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Masood
' Date      : 08/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
63040     On Error GoTo Form_Load_Error


63050     Me.Caption = "NetAcquire - Auto-Generate Comments Microbiology"
63060     ClearValues
63070     Exit Sub


Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

63080     intEL = Erl
63090     strES = Err.Description
63100     LogError "frmAutoGenerateCommentsMicro", "Form_Load", intEL, strES
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearValues
' Author    : Masood
' Date      : 13/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ClearValues()
63110     On Error GoTo ClearValues_Error


63120     FillSites
63130     FillOrgNames
63140     FillWards cmbWard, HospName(0)
63150     cmbWard.AddItem ("Any")
63160     cmbWard = "Any"
63170     cmbSite = "Any"
63180     cmbOrgName = "Select"
63190     FillGrid
63200     txtCounter = ""
63210     txtAgeFromDays = ""
63220     txtAgeToDays = ""
63230     chkPhoneAlert.Value = 0
63240     txtComment = ""
63250     Call EnableControls("C")


63260     dtStart = Date
63270     dtEnd = Date + 365
63280     dtSchDate = Date

63290     Exit Sub

ClearValues_Error:

          Dim strES As String
          Dim intEL As Integer

63300     intEL = Erl
63310     strES = Err.Description
63320     LogError "frmAutoGenerateCommentsMicro", "ClearValues", intEL, strES

End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnableControls
' Author    : Masood
' Date      : 14/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub EnableControls(Operation As String)
63330     On Error GoTo EnableControls_Error

63340     If UCase(Operation) = UCase("N") Then
63350         cmdSaveOrders.Enabled = True
63360         cmdEdit.Enabled = False
63370     ElseIf UCase(Operation) = UCase("E") Then
63380         If txtCounter = "" Then
63390             cmdEdit.Enabled = True
63400         Else
63410             cmdEdit.Enabled = False
63420         End If
63430     ElseIf UCase(Operation) = UCase("Edited") Then
63440         cmdEdit.Enabled = False
63450     ElseIf UCase(Operation) = UCase("C") Then
63460         cmdEdit.Enabled = False
63470     Else
63480         cmdEdit.Enabled = False
63490     End If
63500     cmdSaveOrders.Enabled = True

63510     Exit Sub


EnableControls_Error:

          Dim strES As String
          Dim intEL As Integer

63520     intEL = Erl
63530     strES = Err.Description
63540     LogError "frmAutoGenerateCommentsMicro", "EnableControls", intEL, strES
End Sub



'---------------------------------------------------------------------------------------
' Procedure : FillGrid
' Author    : Masood
' Date      : 08/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub FillGrid()

63550     On Error GoTo FillGrid_Error
          Dim sql As String
          Dim tb As New ADODB.Recordset
          Dim s As String

63560     GridHead

63570     sql = "Select * from MicroAutoCommentAlert  " & _
              " order by ListOrder"
63580     Set tb = New Recordset
63590     RecOpenClient 0, tb, sql
63600     Do While Not tb.EOF

63610         s = tb!Counter & vbTab & tb!OrganismName & vbTab & tb!Site & vbTab & tb!PatientLocation & vbTab & tb!PatientAgeFrom & vbTab & tb!PatientAgeTo & vbTab & tb!DateStart & vbTab & tb!DateEnd & vbTab & tb!Comment & vbTab & IIf((tb!PhoneAlert = True), "True", "False") & vbTab & tb!PhoneAlertDateTime
63620         g.AddItem s
63630         tb.MoveNext
63640     Loop


63650     With g
63660         .ColWidth(0) = 0
63670     End With


63680     Exit Sub


FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

63690     intEL = Erl
63700     strES = Err.Description
63710     LogError "frmAutoGenerateCommentsMicro", "FillGrid", intEL, strES, sql
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GridHead
' Author    : Masood
' Date      : 08/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub GridHead()
63720     On Error GoTo GridHead_Error
          Dim GridFormat As String

63730     With g
63740         .Clear
63750         .Rows = 1
63760         .Cols = 11
63770         GridFormat = "<ListIndex      " & _
                  "|<OrganismName       " & _
                  "|<Site        " & _
                  "|<Pat. Location      " & _
                  "|<Age From " & _
                  "|<Age To    " & _
                  "|<Date Start      " & _
                  "|<DateEnd        " & _
                  "|<Comment                 " & _
                  "|<PhoneAlert    " & _
                  "|<Scheduled Date    "
63780         .FormatString = GridFormat
63790     End With


63800     Exit Sub


GridHead_Error:

          Dim strES As String
          Dim intEL As Integer

63810     intEL = Erl
63820     strES = Err.Description
63830     LogError "frmAutoGenerateCommentsMicro", "GridHead", intEL, strES
End Sub





Private Sub Form_Unload(Cancel As Integer)
    '    End

End Sub

Private Sub g_Click()
          'cmdEdit.Enabled = True
63840     EnableControls ("E")
          '    EditRowSelected
End Sub




'---------------------------------------------------------------------------------------
' Procedure : EditRowSelected
' Author    : Masood
' Date      : 13/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub EditRowSelected()

63850     On Error GoTo EditRowSelected_Error
          Dim RowSel As Integer
63860     With g
63870         RowSel = .RowSel
63880         If RowSel = 0 And cmdEdit.Enabled = True Then
63890             Exit Sub
63900         End If

63910         txtCounter = .TextMatrix(RowSel, 0)    'tb!txtCounter
63920         cmbOrgName = .TextMatrix(RowSel, 1)
63930         cmbSite = IIf(.TextMatrix(RowSel, 2) = "", "Any", .TextMatrix(RowSel, 2))
63940         cmbWard = IIf(.TextMatrix(RowSel, 3) = "", "Any", .TextMatrix(RowSel, 3))
63950         txtAgeFromDays = .TextMatrix(RowSel, 4)
63960         txtAgeToDays = .TextMatrix(RowSel, 5)
63970         dtStart = .TextMatrix(RowSel, 6)
63980         dtEnd = .TextMatrix(RowSel, 7)
63990         txtComment = .TextMatrix(RowSel, 8)



64000         chkPhoneAlert.Value = IIf(.TextMatrix(RowSel, 9) <> "" And (.TextMatrix(RowSel, 9) = True), 1, 0)
64010         If IsDate(.TextMatrix(RowSel, 10)) Then
64020             dtSchDate = .TextMatrix(RowSel, 10)
64030         End If
64040         Call EnableControls("Edited")
64050     End With

64060     Exit Sub


EditRowSelected_Error:

          Dim strES As String
          Dim intEL As Integer

64070     intEL = Erl
64080     strES = Err.Description
64090     LogError "frmAutoGenerateCommentsMicro", "EditRowSelected", intEL, strES
End Sub
