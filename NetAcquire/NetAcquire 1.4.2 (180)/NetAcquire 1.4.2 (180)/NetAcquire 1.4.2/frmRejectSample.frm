VERSION 5.00
Begin VB.Form frmRejectSample 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire -----Reject Sample"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   2520
      TabIndex        =   8
      Top             =   2040
      Width           =   555
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   800
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1845
      Width           =   1275
   End
   Begin VB.CommandButton cmdSaveBio 
      Caption         =   "&Save"
      Height          =   800
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1845
      Width           =   1275
   End
   Begin VB.Frame FrameCoagulation 
      Caption         =   "Coagulation Sample"
      Height          =   1275
      Left            =   3375
      TabIndex        =   3
      Top             =   315
      Width           =   2940
      Begin VB.ComboBox cmbCoag 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Text            =   "cmbCoag"
         Top             =   585
         Width           =   2715
      End
      Begin VB.CheckBox chkCoagReject 
         Caption         =   "Reject Coagulation Sample"
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   315
         Width           =   2490
      End
   End
   Begin VB.Frame frameBio 
      Caption         =   "Biochemistry Sample"
      Height          =   1275
      Left            =   225
      TabIndex        =   0
      Top             =   315
      Width           =   2985
      Begin VB.CheckBox chkBioReject 
         Caption         =   "Reject Biochemistry Sample"
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   315
         Width           =   2490
      End
      Begin VB.ComboBox cmbBio 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Text            =   "cmbBio"
         Top             =   630
         Width           =   2760
      End
   End
End
Attribute VB_Name = "frmRejectSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSavebio_Click()

      Dim sql As String
      Dim tb As Recordset

43070 On Error GoTo cmdSaveBio_Error

43080 frmEditAll.txtSampleID = Format(Val(frmEditAll.txtSampleID))
43090 If Val(frmEditAll.txtSampleID) = 0 Then Exit Sub
43100 If chkBioReject.Value = 1 Then
43110     sql = "IF EXISTS(SELECT * FROM BioResults " & _
                "          WHERE SampleID = @sampleid0 " & _
                "          AND Code = '@Code1' ) " & _
                "  INSERT INTO BioRepeats " & _
                "  (SampleID, Code, Result, Valid, Printed, RunTime, RunDate, " & _
                "   Units, SampleType, Analyser, Faxed, " & _
                "   Healthlink) VALUES " & _
                "  (@sampleid0, '@Code1', '@result2', @valid3, @printed4, @RunTime5, @RunDate6, " & _
                "  '@Units9', '@SampleType10', '@Analyser11', " & _
                "  @Faxed12, @Healthlink18) " & _
                "ELSE " & _
                "  INSERT INTO BioResults " & _
                "  (SampleID, Code, Result, Valid, Printed, RunTime, RunDate, " & _
                "   Units, SampleType, Analyser, Faxed, " & _
                "   Healthlink) VALUES " & _
                "  (@sampleid0, '@Code1', '@result2', @valid3, @printed4, @RunTime5, @RunDate6, " & _
                "  '@Units9', '@SampleType10', '@Analyser11', " & _
                "  @Faxed12, @Healthlink18) "

43120     sql = Replace(sql, "@sampleid0", frmEditAll.txtSampleID)
43130     sql = Replace(sql, "@Code1", Code)
43140     sql = Replace(sql, "@result2", cmbBio.Text)
43150     sql = Replace(sql, "@valid3", 0)
43160     sql = Replace(sql, "@printed4", 0)
43170     sql = Replace(sql, "@RunTime5", Format$(Now, "'dd/mmm/yyyy hh:mm:ss'"))
43180     sql = Replace(sql, "@RunDate6", Format$(Now, "'dd/mmm/yyyy'"))
43190     sql = Replace(sql, "@Units9", cmbUnits)
43200     sql = Replace(sql, "@SampleType10", ListCodeFor("ST", cmbSampleType))
43210     sql = Replace(sql, "@Analyser11", "Manual")
43220     sql = Replace(sql, "@Faxed12", 0)
43230     sql = Replace(sql, "@Healthlink18", 0)

43240     Cnxn(0).Execute sql
43250 ElseIf chkCoagReject.Value = 1 Then
43260     sql = "IF EXISTS(SELECT * FROM CoagResults " & _
                "          WHERE SampleID = @sampleid0 " & _
                "          AND Code = '@Code1' ) " & _
                "  INSERT INTO CoagRepeats " & _
                "  (SampleID, Code, Result, Valid, Printed, RunTime, RunDate, " & _
                "   Units,  Analyser, Faxed, " & _
                "   Healthlink) VALUES " & _
                "  (@sampleid0, '@Code1', '@result2', @valid3, @printed4, @RunTime5, @RunDate6, " & _
                "  '@Units9',  '@Analyser11', " & _
                "  @Faxed12, @Healthlink18) " & _
                "ELSE " & _
                "  INSERT INTO CoagResults " & _
                "  (SampleID, Code, Result, Valid, Printed, RunTime, RunDate, " & _
                "   Units, Analyser, Faxed, " & _
                "   Healthlink) VALUES " & _
                "  (@sampleid0, '@Code1', '@result2', @valid3, @printed4, @RunTime5, @RunDate6, " & _
                "  '@Units9',  '@Analyser11', " & _
                "  @Faxed12, @Healthlink18) "

43270     sql = Replace(sql, "@sampleid0", frmEditAll.txtSampleID)
43280     sql = Replace(sql, "@Code1", Code)
43290     sql = Replace(sql, "@result2", cmbCoag.Text)
43300     sql = Replace(sql, "@valid3", 0)
43310     sql = Replace(sql, "@printed4", 0)
43320     sql = Replace(sql, "@RunTime5", Format$(Now, "'dd/mmm/yyyy hh:mm:ss'"))
43330     sql = Replace(sql, "@RunDate6", Format$(Now, "'dd/mmm/yyyy'"))
43340     sql = Replace(sql, "@Units9", cmbUnits)
43350     sql = Replace(sql, "@Analyser11", "Manual")
43360     sql = Replace(sql, "@Faxed12", 0)
43370     sql = Replace(sql, "@Healthlink18", 0)

43380     Cnxn(0).Execute sql
43390 End If



43400 If Validate Then
43410     sql = "UPDATE BioResults SET " & _
                "Valid = 1, " & _
                " ValidateTime = '" & Format$(Now, "dd/MMM/yyyy HH:mm:ss") & "' ," & _
                "Operator = '" & UserCode & "' " & _
                "WHERE SampleID = '" & txtSampleID & "' " & _
                "AND (COALESCE(Valid, 0) = 0)"
43420     Cnxn(0).Execute sql
43430     BioValBy = UserCode
43440 End If



43450 Exit Sub

cmdSaveBio_Error:

      Dim strES As String
      Dim intEL As Integer

43460 intEL = Erl
43470 strES = Err.Description
43480 LogError "frmRejectSample", "cmdSavebio", intEL, strES, sql

End Sub

Private Sub Form_Load()
43490 FillGenericList cmbBio, "RS"
43500 FillGenericList cmbCoag, "RS"
End Sub
