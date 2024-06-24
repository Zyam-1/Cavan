VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmian 
   Caption         =   "NetAcquire - Semen Analysis History"
   ClientHeight    =   2400
   ClientLeft      =   930
   ClientTop       =   4455
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   10500
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   585
      Left            =   9150
      Picture         =   "frmSemenHistory.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1590
      Width           =   1005
   End
   Begin MSFlexGridLib.MSFlexGrid grdHistory 
      Height          =   1785
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3149
      _Version        =   393216
      Cols            =   8
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Sample ID  |<Sample Date/Time |<Count |<Volume |<Consistency |^Motile Progressive|^Motile Non-Progressive|^Non Motile"
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6510
      TabIndex        =   4
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   6090
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   750
      TabIndex        =   2
      Top             =   60
      Width           =   3045
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String
      Dim SRS As New SemenResults
      Dim SR As SemenResult

53840 On Error GoTo FillG_Error

53850 With grdHistory
53860   .Rows = 2
53870   .AddItem ""
53880   .RemoveItem 1
53890 End With

53900 sql = "Select D.PatName, D.SampleDate " & _
            "from Demographics D where " & _
            "D.Chart = '" & lblChart & "' " & _
            "order by D.SampleID desc"
53910 Set tb = New Recordset
53920 RecOpenServer 0, tb, sql
53930 Do While Not tb.EOF
53940   If Trim$(lblName) = "" Then
53950     lblName = tb!PatName & ""
53960   End If

53970   SRS.Load tb!SampleID
53980   If SRS.Count > 0 Then
53990     Set SR = SRS("SpecimenType")
54000     If Not SR Is Nothing Then
54010       If SR.Result = "Infertility Analysis" Then
54020         s = Trim$(tb!SampleID & "") & vbTab & _
                  Format(tb!SampleDate, "dd/mm/yy hh:mm") & vbTab
        
54030         Set SR = SRS("SemenCount")
54040         If Not SR Is Nothing Then
54050           s = s & SR.Result
54060         End If
54070         s = s & vbTab
        
54080         Set SR = SRS("Volume")
54090         If Not SR Is Nothing Then
54100           s = s & SR.Result
54110         End If
54120         s = s & vbTab
        
54130         Set SR = SRS("Consistency")
54140         If Not SR Is Nothing Then
54150           s = s & SR.Result
54160         End If
54170         s = s & vbTab
        
54180         Set SR = SRS("MotilityPro")
54190         If Not SR Is Nothing Then
54200           s = s & SR.Result
54210         End If
54220         s = s & vbTab
        
54230         Set SR = SRS("MotilityNonPro")
54240         If Not SR Is Nothing Then
54250           s = s & SR.Result
54260         End If
54270         s = s & vbTab
        
54280         Set SR = SRS("MotilityNonMotile")
54290         If Not SR Is Nothing Then
54300           s = s & SR.Result
54310         End If
        
54320         grdHistory.AddItem s
54330       End If
54340     End If
54350   End If
        
54360   tb.MoveNext
54370 Loop

54380 With grdHistory
54390   If .Rows > 2 Then
54400     .RemoveItem 1
54410   End If
54420 End With

54430 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

54440 intEL = Erl
54450 strES = Err.Description
54460 LogError "frmSemenHistory", "FillG", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

54470 Unload Me

End Sub

Private Sub Form_Activate()

54480 FillG

End Sub
