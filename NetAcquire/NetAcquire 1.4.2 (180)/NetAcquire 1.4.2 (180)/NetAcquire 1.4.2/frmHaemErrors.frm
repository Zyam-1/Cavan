VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHaemErrors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   8295
   ClientLeft      =   1950
   ClientTop       =   6015
   ClientWidth     =   12825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstAEM 
      Height          =   2010
      Left            =   6540
      TabIndex        =   3
      Top             =   5040
      Width           =   6135
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   6315
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   885
      Left            =   11460
      Picture         =   "frmHaemErrors.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7260
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   9900
      TabIndex        =   0
      Top             =   2700
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid gFlags 
      Height          =   4485
      Left            =   120
      TabIndex        =   6
      Top             =   540
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   7911
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmHaemErrors.frx":1986
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Extended IPU Flags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   2070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sapphire Flags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9900
      TabIndex        =   4
      Top             =   2340
      Width           =   1605
   End
End
Attribute VB_Name = "frmHaemErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mErrorNumber As Long

Private mSampleID As Long

Public Property Let ErrorNumber(ByVal ErrorNumber As String)

4730      mErrorNumber = Val(ErrorNumber)

End Property

Private Sub FillAEM()

          Dim sql As String
          Dim tb As Recordset
          Dim s() As String
          Dim n As Integer

4740      On Error GoTo FillAEM_Error

4750      lstAEM.Clear

4760      sql = "SELECT AnalyserMessage FROM HaemResults WHERE " & _
              "SampleID = '" & mSampleID & "'"
4770      Set tb = New Recordset
4780      RecOpenServer 0, tb, sql
4790      If Not tb.EOF Then
4800          s = Split(tb!AnalyserMessage & "", "|")
4810          If UBound(s) > -1 Then
4820              For n = 0 To UBound(s)
4830                  lstAEM.AddItem s(n)
4840              Next
4850          End If
4860      End If

4870      Exit Sub

FillAEM_Error:

          Dim strES As String
          Dim intEL As Integer

4880      intEL = Erl
4890      strES = Err.Description
4900      LogError "fHaemErrors", "FillAEM", intEL, strES, sql

End Sub

Private Sub FillList()

          Dim n As Integer
          Dim Trial As Long

4910      On Error GoTo FillList_Error

4920      List1.Clear

          '90          If Mid(m, 91, 1) = "1" Then Errors = 1'abnormal differentiation
          '100         If Mid(m, 92, 1) = "1" Then Errors = Errors + 2'abnormal morphology
          '110         If Mid(m, 93, 1) = "1" Then Errors = Errors + 4'abnormal count
          '120         If Mid(m, 95, 1) = "1" Then Errors = Errors + 8'abnormal result
4930      If HospName(0) = "Mallow" Then
4940          For n = 0 To 3
4950              Trial = 2 ^ n
4960              If mErrorNumber And Trial Then
4970                  List1.AddItem Choose(n + 1, "Abnormal Differentiation", _
                          "Abnormal Morphology", _
                          "Abnormal Count", _
                          "Abnormal Result")
4980              End If
4990          Next
5000      Else
5010          For n = 0 To 25
5020              Trial = 2 ^ n
5030              If mErrorNumber And Trial Then
5040                  List1.AddItem Choose(n + 1, "Moving Average", "DFLT Flag", "Blast Flag", _
                          "Variant Lymph", "DFLT (N)", "DFLT (E)", _
                          "DFLT (L)", "IG Flag", "Band Flag", _
                          "DFLT (M)", "DFLT (B)", "IG/Bands", _
                          "FWBC", "WBC Count", "NRBC", _
                          "DLTA", "NWBC", "RBC Morph", _
                          "RRBC", "Plt Recount", "LRI", _
                          "URI", "NOC Flow", "WOC Flow", _
                          "RBC Flow", "Sampling Error")
5050              End If
5060          Next
5070      End If

5080      Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

5090      intEL = Erl
5100      strES = Err.Description
5110      LogError "fHaemErrors", "FillList", intEL, strES

End Sub

Private Function FillProcessingErrors() As Boolean
          'returns true if SampleProcessingError contains "^"
          'only contains "^" if result comes from Sapphire

          Dim X As Long
          Dim MSP As Long
          Dim LSP As Long
          Dim n As Integer
          Dim Trial As Long
          Dim MorphError As Long
          Dim SapphireErrors() As String

5120      On Error GoTo FillProcessingErrors_Error

5130      FillProcessingErrors = False

5140      List1.Clear
5150      List2.Clear

          'C|2|I|DI^00648^00002|I<CR>
          'Interpretation of MSP and LSP, “00648”,“00002,” yields the decimal numbers 648 and 2,
          'which result in a bit pattern of
          '0000 0010 1000 1000 0000 0000 0000 0010
          'The bit pattern can be used in a binary test for
          'data invalidation or interpreted against Table 17
          'to indicate the following sample processing
          'invalidations:
          'Bit 1 = Clog detected during impedance
          'measurement
          'Bit 19 = Reticulocyte Reagent not detected
          'Bit 23 = Retic Count Rate error
          'Bit 25 = RBCi Count Rate error

5160      If InStr(SampleProcessingError, "^") = 0 Then Exit Function

5170      FillProcessingErrors = True

5180      SapphireErrors = Split(SampleProcessingError, "^")
5190      MSP = Val(SapphireErrors(0)) * 65536
5200      LSP = Val(SapphireErrors(1))
5210      X = MSP + LSP
5220      If UBound(SapphireErrors) > 1 Then
5230          MorphError = Val(SapphireErrors(2))
5240      Else
5250          MorphError = 0
5260      End If

5270      If X <> 0 Then
5280          For n = 0 To 29
5290              Trial = 2 ^ n
5300              If X And Trial Then
5310                  List2.AddItem Choose(n + 1, "Clot detected during aspiration", "Clog detected during impedance measurement", _
                          "Residual Fluid Detected in WBC Dilution Cup", "Laser Fault", _
                          "WBC dilution cup temperature out of range", "WBC dilution cup motor voltage out of range", _
                          "Autoloader/Mixhead Failure", "Hemoglobin reagent syringe failed to home", _
                          "Missing reagent tube, assay not performed", "Incorrect reagent tube, assay not performed", _
                          "Diluent/sheath reagent syringe failed to home", "RETC Reagent Syringe failed to home", _
                          "WBC Reagent Syringes failed to home", "Pneumatics Failure in high pressure range", _
                          "Pneumatics Failure in medium pressure range", "Pneumatics Failure in low pressure range", _
                          "Pneumatics Failure in vacuum range", "Sample aspiration pump failed to home", _
                          "Vent/aspirate assembly failure", "Reticulocyte reagent not detected", "This Bit # is not assigned", _
                          "Insufficient reagent to complete multi-tube assay", "Short Sample", "RETC Count Rate error", _
                          "WBC Count Rate error", "RBCi Count Rate error", "PLTi Count Rate error", "PLTo Count Rate error", _
                          "CD61 Optical Platelet Stream Mismatch", "CD3/4/8 Count Rate/Mismatch")
5320              End If
5330          Next

5340      End If

5350      If MorphError <> 0 Then
5360          For n = 3 To 12
5370              Trial = 2 ^ n
5380              If MorphError And Trial Then
5390                  List1.AddItem Choose(n - 2, "Nonviable WBCs", _
                          "Bands", _
                          "Blasts", _
                          "Variant Lymphoids", _
                          "Immature Granulocytes", _
                          "Unidentified Fluorescent Cells", _
                          "Resistant RBCs", _
                          "Asymmetric RBC", _
                          "High Number Immature Reticulocytes", _
                          "Platelet Clumps")
5400              End If
5410          Next
5420      End If

          'Bit Values
          '8 "4C|3|I|SP^Nonviable WBCs^NO^0.00|I<CR>"
          '16 "4C|4|I|SP^Bands^YES^0.90|I<CR>"
          '32 "4C|5|I|SP^Blasts^NO^0.00|I<CR>
          '64 "4C|6|I|SP^Variant Lymphoids^NO^0.00|I<CR>
          '128 "4C|7|I|SP^Immature Granulocytes^YES^0.60|I<CR>
          '256 "4C|8|I|SP^Unidentified Fluorescent Cells^NO^0.00|I<CR>
          '512 "4C|9|I|SP^Resistant RBCs^YES^0.00|I<CR>
          '1024 "4C|10|I|SP^Asymmetric RBC^NO^0.00^|I<CR>
          '2048 "4C|11|I|SP^High Number Immature Reticulocytes^NO^0.00|I<CR>
          '4096 "4C|12|I|SP^Platelet Clumps^NO^0.00|I<CR>

5430      Exit Function

FillProcessingErrors_Error:

          Dim strES As String
          Dim intEL As Integer

5440      intEL = Erl
5450      strES = Err.Description
5460      LogError "fHaemErrors", "FillProcessingErrors", intEL, strES

End Function

Private Function FillgFlags()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

5470      On Error GoTo FillgFlags_Error

5480      sql = "SELECT * FROM HaemFlags WHERE ( UPPER(FlagType) <> UPPER('Hematology') OR UPPER(Flag) like 'ACTION%' ) " & _
              " and   SampleID = " & mSampleID
5490      Set tb = New Recordset
5500      RecOpenServer 0, tb, sql
5510      While Not tb.EOF
5520          s = tb!DateTime & "" & vbTab & _
                  tb!FlagType & "" & vbTab & _
                  tb!Flag & ""

5530          gFlags.AddItem s
5540          tb.MoveNext
5550      Wend
5560      If gFlags.Rows > 2 Then
5570          gFlags.RemoveItem 1
5580      End If

5590      Exit Function

FillgFlags_Error:

          Dim strES As String
          Dim intEL As Integer

5600      intEL = Erl
5610      strES = Err.Description
5620      LogError "frmHaemErrors", "FillgFlags", intEL, strES, sql

End Function


Private Sub bcancel_Click()

5630      Unload Me

End Sub


Private Sub Form_Activate()

5640      If FillProcessingErrors() = False Then
5650          FillList
5660      End If

5670      FillAEM
5680      FillgFlags

End Sub

Public Property Let SampleID(ByVal lNewValue As Long)

5690      mSampleID = lNewValue

End Property
