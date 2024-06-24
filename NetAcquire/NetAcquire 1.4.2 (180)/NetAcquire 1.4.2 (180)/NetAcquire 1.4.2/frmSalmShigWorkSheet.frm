VERSION 5.00
Begin VB.Form frmSalmShigWorkSheet 
   Caption         =   "NetAcquire - Salmonella/Shigella"
   ClientHeight    =   4305
   ClientLeft      =   105
   ClientTop       =   420
   ClientWidth     =   12210
   LinkTopic       =   "Form2"
   ScaleHeight     =   4305
   ScaleWidth      =   12210
   Begin VB.Frame Frame4 
      Height          =   1275
      Index           =   1
      Left            =   270
      TabIndex        =   56
      Top             =   150
      Width           =   7275
      Begin VB.Label lblSex 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5940
         TabIndex        =   66
         Top             =   330
         Width           =   705
      End
      Begin VB.Label lblAge 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4860
         TabIndex        =   65
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblDoB 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2850
         TabIndex        =   64
         Top             =   330
         Width           =   1515
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   690
         TabIndex        =   63
         Top             =   780
         Width           =   5955
      End
      Begin VB.Label lblChart 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   660
         TabIndex        =   62
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Chart #"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   61
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   60
         Top             =   810
         Width           =   420
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   0
         Left            =   2370
         TabIndex        =   59
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   0
         Left            =   4530
         TabIndex        =   58
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   0
         Left            =   5610
         TabIndex        =   57
         Top             =   360
         Width           =   270
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   10260
      Picture         =   "frmSalmShigWorkSheet.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   480
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   735
      Left            =   8880
      Picture         =   "frmSalmShigWorkSheet.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   480
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      Caption         =   "Salmonella"
      Height          =   2595
      Left            =   240
      TabIndex        =   24
      Top             =   1500
      Width           =   7275
      Begin VB.CheckBox chkColindale 
         Alignment       =   1  'Right Justify
         Caption         =   "Colindale"
         Height          =   195
         Left            =   3540
         TabIndex        =   39
         Top             =   2040
         Width           =   945
      End
      Begin VB.CheckBox chkB17 
         Caption         =   "Check4"
         Height          =   255
         Left            =   4590
         TabIndex        =   38
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chkB15 
         Caption         =   "Check3"
         Height          =   255
         Left            =   4290
         TabIndex        =   37
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chkB16 
         Caption         =   "Check2"
         Height          =   255
         Left            =   4590
         TabIndex        =   36
         Top             =   1290
         Width           =   255
      End
      Begin VB.CheckBox chkB12 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4290
         TabIndex        =   35
         Top             =   1290
         Width           =   255
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Check1"
         Height          =   195
         Left            =   5670
         TabIndex        =   34
         Top             =   870
         Width           =   225
      End
      Begin VB.CheckBox chkLittleI 
         Caption         =   "Check1"
         Height          =   225
         Left            =   5310
         TabIndex        =   33
         Top             =   840
         Width           =   225
      End
      Begin VB.CommandButton bSens 
         Caption         =   "Sensitivity"
         Height          =   315
         Index           =   0
         Left            =   5970
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1500
         Width           =   1125
      End
      Begin VB.CheckBox chkR 
         Alignment       =   1  'Right Justify
         Caption         =   "Rapid 1"
         Height          =   195
         Index           =   1
         Left            =   3930
         TabIndex        =   31
         Top             =   450
         Width           =   855
      End
      Begin VB.CheckBox chkR 
         Alignment       =   1  'Right Justify
         Caption         =   "Rapid 3"
         Height          =   195
         Index           =   3
         Left            =   3930
         TabIndex        =   30
         Top             =   870
         Width           =   855
      End
      Begin VB.CheckBox chkR 
         Alignment       =   1  'Right Justify
         Caption         =   "Rapid 2"
         Height          =   195
         Index           =   2
         Left            =   3930
         TabIndex        =   29
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox txtHAntigen 
         Height          =   285
         Left            =   5190
         TabIndex        =   28
         Top             =   450
         Width           =   1245
      End
      Begin VB.TextBox txtSalmType 
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtSalmID 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   26
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtColindaleResult 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   4530
         MaxLength       =   40
         TabIndex        =   25
         Top             =   2010
         Width           =   2625
      End
      Begin VB.Label lblPolyH2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   53
         Top             =   1590
         Width           =   1275
      End
      Begin VB.Label lblPolyH 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   52
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label lblPolyO 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Positive"
         Height          =   255
         Left            =   2040
         TabIndex        =   51
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "H Antigen"
         Height          =   195
         Index           =   0
         Left            =   5220
         TabIndex        =   50
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "1,7"
         Height          =   195
         Left            =   4860
         TabIndex        =   49
         Top             =   1590
         Width           =   225
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "1,6"
         Height          =   195
         Left            =   4860
         TabIndex        =   48
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "1,5"
         Height          =   195
         Left            =   3960
         TabIndex        =   47
         Top             =   1590
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "1,2"
         Height          =   195
         Left            =   3960
         TabIndex        =   46
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Poly H Phase 2"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   45
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Poly H Phase 1 && 2"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   1230
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Polyvalent-O Groups A-S"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   390
         Width           =   1785
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Other"
         Height          =   195
         Left            =   5910
         TabIndex        =   42
         Top             =   870
         Width           =   390
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "i"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   5190
         TabIndex        =   41
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Index           =   0
         Left            =   1590
         TabIndex        =   40
         Top             =   690
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Shigella"
      Height          =   2595
      Index           =   0
      Left            =   7590
      TabIndex        =   0
      Top             =   1500
      Width           =   4485
      Begin VB.CheckBox chkSonn 
         Alignment       =   1  'Right Justify
         Caption         =   "Phase 2"
         Height          =   195
         Index           =   1
         Left            =   3390
         TabIndex        =   17
         Top             =   1710
         Width           =   885
      End
      Begin VB.CheckBox chkSonn 
         Alignment       =   1  'Right Justify
         Caption         =   "Phase 1"
         Height          =   195
         Index           =   0
         Left            =   3390
         TabIndex        =   16
         Top             =   1470
         Width           =   885
      End
      Begin VB.CommandButton bSens 
         Caption         =   "Sensitivity"
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   15
         Top             =   1500
         Width           =   975
      End
      Begin VB.CheckBox chkFlex 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         Height          =   225
         Index           =   7
         Left            =   240
         TabIndex        =   14
         Top             =   1500
         Width           =   405
      End
      Begin VB.CheckBox chkFlex 
         Caption         =   "Y"
         Height          =   225
         Index           =   8
         Left            =   720
         TabIndex        =   13
         Top             =   1500
         Width           =   465
      End
      Begin VB.CheckBox chkFlex 
         Caption         =   "6"
         Height          =   225
         Index           =   6
         Left            =   720
         TabIndex        =   12
         Top             =   1230
         Width           =   465
      End
      Begin VB.CheckBox chkFlex 
         Caption         =   "5"
         Height          =   225
         Index           =   5
         Left            =   720
         TabIndex        =   11
         Top             =   960
         Width           =   465
      End
      Begin VB.CheckBox chkFlex 
         Caption         =   "4"
         Height          =   225
         Index           =   4
         Left            =   720
         TabIndex        =   10
         Top             =   690
         Width           =   465
      End
      Begin VB.CheckBox chkFlex 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1230
         Width           =   405
      End
      Begin VB.CheckBox chkFlex 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   405
      End
      Begin VB.CheckBox chkFlex 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   690
         Width           =   405
      End
      Begin VB.CheckBox chkFlex 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly"
         Height          =   225
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   420
         Width           =   615
      End
      Begin VB.CheckBox chkDys 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly 3-10"
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   5
         Top             =   900
         Width           =   975
      End
      Begin VB.CheckBox chkDys 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly 1-10"
         Height          =   255
         Index           =   0
         Left            =   3300
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkBoy 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly 12-15"
         Height          =   195
         Index           =   2
         Left            =   1590
         TabIndex        =   3
         Top             =   1110
         Width           =   1065
      End
      Begin VB.CheckBox chkBoy 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly 7-11"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   2
         Top             =   870
         Width           =   975
      End
      Begin VB.CheckBox chkBoy 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly 1-6"
         Height          =   195
         Index           =   0
         Left            =   1770
         TabIndex        =   1
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "S Dysenteriae"
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
         Index           =   2
         Left            =   3090
         TabIndex        =   23
         Top             =   390
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "S Sonnei"
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
         Index           =   1
         Left            =   3480
         TabIndex        =   22
         Top             =   1260
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "S Boydii"
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
         Index           =   2
         Left            =   1920
         TabIndex        =   21
         Top             =   390
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "S Flexneri"
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
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   210
         Width           =   855
      End
      Begin VB.Label lblShigellaType 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   270
         TabIndex        =   19
         Top             =   2040
         Width           =   4005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   18
         Top             =   1800
         Width           =   435
      End
   End
   Begin VB.Label lblSampleID 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   9360
      TabIndex        =   68
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   8580
      TabIndex        =   67
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "frmSalmShigWorkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SaveSalmShig()

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim Counter As Integer
Dim SIDOff As Long

SIDOff = Val(lblSampleID) + sysOptMicroOffset(0)

sql = "Select * from SalmShig where " & _
      "SampleID = '" & SIDOff & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then tb.AddNew

tb!SampleID = SIDOff

tb!PolyO = Trim$(Left(lblPolyO & " ", 1))
tb!PolyH = Trim$(Left(lblPolyH & " ", 1))
tb!PolyH2 = Trim$(Left(lblPolyH2 & " ", 1))

tb!SalmType = txtSalmType


Counter = 0
If chkR(1) Then Counter = 1
If chkR(2) Then Counter = Counter + 2
If chkR(3) Then Counter = Counter + 4
tb!Rapid = Counter

tb!LittleI = chkLittleI
tb!Other = chkOther

Counter = 0
If chkB12 Then Counter = 1
If chkB15 Then Counter = Counter + 2
If chkB16 Then Counter = Counter + 4
If chkB17 Then Counter = Counter + 8
tb!b12 = Counter

tb!SalmIdent = txtSalmID
tb!Colindale = chkColindale
tb!ColindaleResult = txtColindaleResult

tb!ShigType = lblShigellaType

Counter = 0
For n = 0 To 8
  If chkFlex(n) Then Counter = Counter + 2 ^ n
Next
tb!Flex = Counter

Counter = 0
For n = 0 To 2
  If chkBoy(n) Then Counter = Counter + 2 ^ n
Next
tb!Boy = Counter

Counter = 0
For n = 0 To 1
  If chkDys(n) Then Counter = Counter + 2 ^ n
Next
tb!Dys = Counter

Counter = 0
For n = 0 To 1
  If chkSonn(n) Then Counter = Counter + 2 ^ n
Next
tb!Sonn = Counter

tb.Update

End Sub

Private Sub chkB12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSave.Enabled = True

End Sub


Private Sub chkB15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSave.Enabled = True

End Sub


Private Sub chkB16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSave.Enabled = True

End Sub


Private Sub chkB17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSave.Enabled = True

End Sub


Private Sub chkColindale_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSave.Enabled = True

End Sub


Private Sub chkLittleI_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSave.Enabled = True

End Sub


Private Sub chkOther_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdSave.Enabled = True

End Sub


Private Sub chkR_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Pattern As String

Pattern = IIf(chkR(1) = 1, "+", "-") & _
          IIf(chkR(2) = 1, "+", "-") & _
          IIf(chkR(3) = 1, "+", "-")

Select Case Pattern
  Case "++-": txtHAntigen = "b"
  Case "+-+": txtHAntigen = "d"
  Case "+++": txtHAntigen = "E Complex"
  Case "--+": txtHAntigen = "G Complex"
  Case "-++": txtHAntigen = "k"
  Case "-+-": txtHAntigen = "L Complex"
  Case "+--": txtHAntigen = "r"
  Case "---": txtHAntigen = ""
End Select

cmdSave.Enabled = True

End Sub

Private Sub cmdCancel_Click()

If cmdSave.Enabled Then
  If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
  End If
End If

Unload Me

End Sub

Private Sub chkBoy_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim n As Integer

lblShigellaType = ""

For n = 0 To 2
  If chkBoy(n) = 1 Then lblShigellaType = "Shigella Boydii"
Next

cmdSave.Enabled = True

End Sub


Private Sub chkFlex_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim n As Integer
Dim Found As Integer

If Index > 0 And Index < 7 Then
  Found = Index
  For n = 1 To 6
    chkFlex(n) = 0
  Next
  chkFlex(Found) = 1
End If

lblShigellaType = ""
Found = False

If chkFlex(0) = 1 Then
  For n = 1 To 6
    If chkFlex(n) = 1 Then
      lblShigellaType = "Shigella Flexneri Type " & n
      Found = True
    End If
  Next
  If Found Then
    If chkFlex(7) = 1 And chkFlex(8) = 0 Then
      lblShigellaType = lblShigellaType & " Variant X"
    ElseIf chkFlex(8) = 1 And chkFlex(7) = 0 Then
      lblShigellaType = lblShigellaType & " Variant Y"
    End If
  Else
    lblShigellaType = ""
  End If
End If

cmdSave.Enabled = True

End Sub


Private Sub chkdys_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim n As Integer

lblShigellaType = ""

For n = 0 To 1
  If chkDys(n) = 1 Then lblShigellaType = "Shigella Dysenteriae"
Next

cmdSave.Enabled = True

End Sub


Private Sub LoadSalmShig()

Dim tb As Recordset
Dim sql As String
Dim n As Integer
Dim SIDOff As Long

ClearSalmShig

SIDOff = Val(lblSampleID) + sysOptMicroOffset(0)

sql = "Select * from SalmShig where " & _
      "SampleID = '" & SIDOff & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  If tb!PolyO & "" = "N" Then
    lblPolyO = "Negative"
    lblPolyO.BackColor = vbGreen
  ElseIf tb!PolyO & "" = "P" Then
    lblPolyO = "Positive"
    lblPolyO.BackColor = vbRed
  End If
  txtSalmType = tb!SalmType & ""
  If tb!PolyH & "" = "N" Then
    lblPolyH = "Negative"
    lblPolyH.BackColor = vbGreen
  ElseIf tb!PolyH & "" = "P" Then
    lblPolyH = "Positive"
    lblPolyH.BackColor = vbRed
  ElseIf tb!PolyH & "" = "I" Then
    lblPolyH = "Indeterminate"
    lblPolyH.BackColor = vbYellow
  End If
  If tb!PolyH2 & "" = "N" Then
    lblPolyH2 = "Negative"
    lblPolyH2.BackColor = vbGreen
  ElseIf tb!PolyH2 & "" = "P" Then
    lblPolyH2 = "Positive"
    lblPolyH2.BackColor = vbRed
  ElseIf tb!PolyH2 & "" = "I" Then
    lblPolyH2 = "Indeterminate"
    lblPolyH2.BackColor = vbYellow
  End If
  
  If Not IsNull(tb!Rapid) Then
    chkR(1) = IIf(tb!Rapid And 1, 1, 0)
    chkR(2) = IIf(tb!Rapid And 2, 1, 0)
    chkR(3) = IIf(tb!Rapid And 4, 1, 0)
  End If
  
  chkLittleI = IIf(tb!LittleI, 1, 0)
  chkOther = IIf(tb!Other, 1, 0)
  If Not IsNull(tb!b12) Then
    chkB12 = IIf(tb!b12 And 1, 1, 0)
    chkB15 = IIf(tb!b12 And 2, 1, 0)
    chkB16 = IIf(tb!b12 And 4, 1, 0)
    chkB17 = IIf(tb!b12 And 8, 1, 0)
  End If
  txtSalmID = tb!SalmIdent & ""
  chkColindale = IIf(tb!Colindale, 1, 0)
  txtColindaleResult = tb!ColindaleResult & ""

  lblShigellaType = tb!ShigType & ""
  For n = 0 To 8
    chkFlex(n) = IIf(tb!Flex And 2 ^ n, 1, 0)
  Next
  For n = 0 To 2
    chkBoy(n) = IIf(tb!Boy And 2 ^ n, 1, 0)
  Next
  For n = 0 To 1
    chkDys(n) = IIf(tb!Dys And 2 ^ n, 1, 0)
    chkSonn(n) = IIf(tb!Sonn And 2 ^ n, 1, 0)
  Next
End If

End Sub

Private Sub ClearSalmShig()
  
Dim n As Integer
  
lblPolyO = ""
lblPolyO.BackColor = &H8000000F

txtSalmType = ""
lblPolyH = ""
lblPolyH.BackColor = &H8000000F

chkR(1) = False
chkR(2) = False
chkR(3) = False
chkLittleI = False
chkOther = False
chkB12 = 0
chkB15 = 0
chkB16 = 0
chkB17 = 0
chkColindale = False
txtColindaleResult = ""
txtSalmID = ""

lblShigellaType = ""
For n = 0 To 8
  chkFlex(n) = False
Next
For n = 0 To 2
  chkBoy(n) = False
Next
For n = 0 To 1
  chkDys(n) = False
  chkSonn(n) = False
Next

End Sub

Private Sub cmdsave_Click()

SaveSalmShig

Unload Me

End Sub

Private Sub Form_Activate()

LoadSalmShig

End Sub

Private Sub Form_Unload(Cancel As Integer)

With frmEditMicrobiology
  .lblSalmonella = txtSalmID
  .lblShigella = lblShigellaType
  .lblColindale = txtColindaleResult
End With

End Sub


Private Sub lblPolyO_Click()

With lblPolyO
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSave.Enabled = True

End Sub


Private Sub lblPolyH_Click()

With lblPolyH
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = "Indeterminate"
    .BackColor = vbYellow
  Case "Indeterminate"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSave.Enabled = True

End Sub



Private Sub lblPolyH2_Click()

With lblPolyH2
  Select Case .Caption
  Case ""
    .Caption = "Negative"
    .BackColor = vbGreen
  Case "Negative"
    .Caption = "Positive"
    .BackColor = vbRed
  Case "Positive"
    .Caption = "Indeterminate"
    .BackColor = vbYellow
  Case "Indeterminate"
    .Caption = ""
    .BackColor = &H8000000F
  End Select
End With

cmdSave.Enabled = True

End Sub


Private Sub chkSonn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim n As Integer

lblShigellaType = ""

For n = 0 To 1
  If chkSonn(n) = 1 Then lblShigellaType = "Shigella Sonnei"
Next

cmdSave.Enabled = True

End Sub


Private Sub lblShigellaType_DblClick()

If lblShigellaType = "" Then
  lblShigellaType = "No Shigella Isolated"
Else
  lblShigellaType = ""
End If

cmdSave.Enabled = True

End Sub


Private Sub txtColindaleResult_KeyPress(KeyAscii As Integer)

cmdSave.Enabled = True

End Sub


Private Sub txtSalmID_DblClick()

txtSalmID = "No Salmonella Isolated"

cmdSave.Enabled = True

End Sub


Private Sub txtSalmType_KeyPress(KeyAscii As Integer)

cmdSave.Enabled = True

End Sub


