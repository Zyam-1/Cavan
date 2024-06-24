VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBatchProductEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Batch Entry"
   ClientHeight    =   4152
   ClientLeft      =   492
   ClientTop       =   528
   ClientWidth     =   8280
   Icon            =   "frmBatchProductEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4152
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox cAuto 
      Caption         =   "Auto-Entry is ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   1200
      TabIndex        =   19
      Top             =   3270
      Value           =   1  'Checked
      Width           =   3045
   End
   Begin VB.TextBox tConc 
      Height          =   285
      Left            =   5310
      TabIndex        =   23
      Text            =   "4.5"
      Top             =   1245
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox tInput 
      Height          =   285
      Left            =   9390
      TabIndex        =   22
      Top             =   240
      Width           =   1875
   End
   Begin VB.ComboBox cGroup 
      Height          =   315
      Left            =   3090
      TabIndex        =   18
      Text            =   "cGroup"
      Top             =   1230
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Height          =   885
      Left            =   3090
      TabIndex        =   13
      Top             =   1620
      Width           =   2745
      Begin VB.TextBox txtReceived 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Text            =   "5"
         Top             =   240
         Width           =   795
      End
      Begin MSComCtl2.UpDown udReceived 
         Height          =   225
         Left            =   1680
         TabIndex        =   16
         Top             =   510
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   402
         _Version        =   393216
         Value           =   5
         BuddyControl    =   "txtReceived"
         BuddyDispid     =   196614
         OrigLeft        =   4410
         OrigTop         =   1170
         OrigRight       =   5175
         OrigBottom      =   1410
         Max             =   1000
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Number of Units Received"
         Height          =   405
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   1245
      End
   End
   Begin MSComCtl2.DTPicker dtReceived 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Top             =   2190
      Width           =   1365
      _ExtentX        =   2413
      _ExtentY        =   550
      _Version        =   393216
      Format          =   158138369
      CurrentDate     =   37294
   End
   Begin MSComCtl2.DTPicker dtExpiry 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Top             =   1710
      Width           =   1365
      _ExtentX        =   2413
      _ExtentY        =   550
      _Version        =   393216
      Format          =   158138369
      CurrentDate     =   37294
   End
   Begin VB.ComboBox cmbProduct 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Text            =   "cmbProduct"
      Top             =   270
      Width           =   4575
   End
   Begin VB.TextBox tVolume 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Text            =   "200"
      Top             =   1245
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   1065
      Left            =   6600
      Picture         =   "frmBatchProductEntry.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "cmdSave"
      Top             =   1230
      Width           =   1245
   End
   Begin VB.TextBox txtBatchNumber 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   780
      Width           =   4305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   6600
      Picture         =   "frmBatchProductEntry.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "cmdCancel"
      Top             =   2670
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   330
      TabIndex        =   26
      Top             =   2730
      Width           =   5505
      _ExtentX        =   9716
      _ExtentY        =   402
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lConcP 
      AutoSize        =   -1  'True
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5910
      TabIndex        =   25
      Top             =   1290
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lConc 
      AutoSize        =   -1  'True
      Caption         =   "Concentration"
      Height          =   195
      Left            =   4260
      TabIndex        =   24
      Top             =   1290
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tInput ->"
      Height          =   255
      Left            =   8550
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblReadError 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entry Error"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4290
      TabIndex        =   20
      Top             =   3330
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   2640
      TabIndex        =   17
      Top             =   1290
      Width           =   435
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      Height          =   195
      Left            =   960
      TabIndex        =   10
      Top             =   330
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "ml"
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   1290
      Width           =   150
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Unit Volume"
      Height          =   195
      Left            =   630
      TabIndex        =   6
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Batch Number"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Expiry Date"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   1770
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date Received"
      Height          =   195
      Left            =   450
      TabIndex        =   1
      Top             =   2250
      Width           =   1080
   End
End
Attribute VB_Name = "frmBatchProductEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EnableControls(ByVal Enable As Boolean)
        
10    cmbProduct.Enabled = Enable
20    txtBatchNumber.Enabled = Enable
30    cGroup.Enabled = Enable
40    dtExpiry.Enabled = Enable

End Sub

Private Sub CheckIfAlb()

10    If cmbProduct = "Albumin" Then
20      tConc.Visible = True
30      lConc.Visible = True
40      lConcP.Visible = True
50    Else
60      tConc.Visible = False
70      lConc.Visible = False
80      lConcP.Visible = False
90    End If

End Sub
Private Sub FillcmbProduct()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillcmbProduct_Error

20    sql = "Select * from Lists where " & _
            "ListType = 'B' " & _
            "Order by ListOrder"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    cmbProduct.Clear
60    Do While Not tb.EOF
70      cmbProduct.AddItem tb!Text & ""
80      tb.MoveNext
90    Loop

100   Exit Sub

FillcmbProduct_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmBatchProductEntry", "FillcmbProduct", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdSave_Click()
        
      Dim BPs As New BatchProducts

10    On Error GoTo cmdSave_Click_Error

20    If Trim$(txtBatchNumber) = "" Then
30      iMsg "Batch Number?", vbQuestion
40      If TimedOut Then Unload Me: Exit Sub
50      Exit Sub
60    End If

70    If DateDiff("d", Now, dtExpiry) < 0 Then
80      iMsg "Expired Expiry Date entered!" & vbCrLf & _
             "Check Expiry Date." & dtExpiry & vbCrLf & _
             "Batch cannot be entered when expired.", vbCritical
90      If TimedOut Then Unload Me: Exit Sub
100     Exit Sub
110   End If

120   If Val(tVolume) = 0 Then
130     iMsg "Unit Volume?", vbQuestion
140     If TimedOut Then Unload Me: Exit Sub
150     Exit Sub
160   End If

170   If UCase$(cmbProduct) <> "UNIPLAS" Then
180     If InStr(UCase$(cmbProduct), "PLAS") <> 0 And cGroup = "" Then
190       iMsg "Enter Group", vbExclamation
200       If TimedOut Then Unload Me: Exit Sub
210       Exit Sub
220     End If
230   End If

240   If Len(txtBatchNumber) <> 9 And InStr(UCase$(cmbProduct), "ALB") = 0 Then
250     Answer = iMsg("Batch Numbers are normally 9 digits." & vbCrLf & _
                "You have entered " & txtBatchNumber & _
                " (" & Format$(Len(txtBatchNumber)) & " digits.)" & vbCrLf & _
                "Is this correct?", vbQuestion + vbYesNo, "Batch Entry")
260     If TimedOut Then Unload Me: Exit Sub
270     If Answer = vbNo Then
280       Exit Sub
290     End If
300     LogReasonWhy "Batch number " & txtBatchNumber & " entered after warning.", "Batch"
310   End If

320   BPs.LoadSpecificBatch txtBatchNumber
330   If BPs.Count > 0 Then
340     If Format$(dtExpiry, "dd/MMM/yyyy") <> Format$(BPs.Item(1).DateExpiry, "dd/MMM/yyyy") Then
350       iMsg "Historic Expiry Date is " & Format$(BPs.Item(1).DateExpiry, "dd/MMM/yyyy") & vbCrLf & _
               "Cannot change Expiry Date here." & vbCrLf & vbCrLf & _
               "(To change Expiry Date go to" & vbCrLf & _
               "Batches/Amendment)"
360       If TimedOut Then Unload Me: Exit Sub
370       Exit Sub
380     End If
390   End If

400   SaveToBatchProducts

410   txtBatchNumber = ""
420   txtReceived = ""

430   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

440   intEL = Erl
450   strES = Err.Description
460   LogError "frmBatchProductEntry", "cmdSave_Click", intEL, strES

End Sub




Private Sub cAuto_Click()

10    If cAuto.Value Then
20      cAuto.Caption = "Auto-Entry is ON"
30      cAuto.ForeColor = vbGreen
40      EnableControls False
50      tInput = ""
60      tInput.SetFocus
70    Else
80      cAuto.Caption = "Auto-Entry is OFF"
90      cAuto.ForeColor = vbRed
100     EnableControls True
110   End If

End Sub

Private Sub cGroup_GotFocus()

10    cAuto.Caption = "Auto-Entry is OFF"
20    cAuto.ForeColor = vbRed
30    cAuto.Value = False

End Sub




Private Sub cmbProduct_Click()

10    CheckIfAlb

End Sub

Private Sub cmbProduct_GotFocus()

10    cAuto.Caption = "Auto-Entry is OFF"
20    cAuto.ForeColor = vbRed
30    cAuto.Value = False

End Sub


Private Sub cmbProduct_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub


Private Sub SaveToBatchProducts()

      Dim ThisIndex As String
      Dim IDX As String
      Dim yy As String
      Dim Rpt As Integer
      Dim BPs As New BatchProducts

      Dim BP As BatchProduct
      Dim C As String

10    On Error GoTo SaveToBatchProducts_Error

20    yy = Mid$(Format$(Now, "YY"), 2)

30    For Rpt = 1 To Val(txtReceived)
40      IDX = GetOptionSetting("BatchIdentifier", "1")
50      SaveOptionSetting "BatchIdentifier", Val(IDX) + 1
60      ThisIndex = yy & Format$(IDX, "0000000")
70      C = CheckCharacterForBatch(ThisIndex)
80      ThisIndex = ThisIndex & C
        
90      Set BP = New BatchProduct
100     With BP
110       .BatchNumber = txtBatchNumber
120       .Identifier = ThisIndex
130       .Product = cmbProduct
140       .UnitVolume = tVolume
150       .DateExpiry = Format(dtExpiry, "dd/MMM/yyyy")
160       .DateReceived = Format(dtReceived, "dd/MMM/yyyy")
170       .UnitGroup = cGroup
180       If tConc.Visible Then
190         .Concentration = tConc
200       Else
210         .Concentration = ""
220       End If
230       .UserName = UserName
240       .EventCode = "C"
250       .RecordDateTime = Format$(Now, "dd/MMM/yyyy HH:nn:ss")
260     End With
270     BPs.Update BP
        
280     PrintBatchIdentifier txtBatchNumber, ThisIndex, cmbProduct, cGroup



      Dim Generic As String
          Dim MSG As udtRS
290       With MSG
300           .UnitNumber = ThisIndex
310           .ProductCode = cmbProduct 'ProductBarCodeFor(lstproduct)
320           Generic = ""
330           If Generic = "Platelets" Then
340               .StorageLocation = strBTCourier_StorageLocation_RoomTempIssueFridge
350           Else
360               .StorageLocation = strBTCourier_StorageLocation_HemoSafeFridge
370           End If
380           .UnitExpiryDate = Format(dtExpiry, "dd/MMM/yyyy")
390           .UnitGroup = cGroup
              'No real Patient details when unit received
400           .StockComment = ""
410           .Chart = ""
420           .PatientHealthServiceNumber = ""
430           .ForeName = ""
440           .SurName = ""
450           .DoB = ""
460           .Sex = ""
470           .PatientGroup = ""
480           .DeReservationDateTime = Format(dtExpiry, "dd-MMM-yyyy hh:mm:ss")
490           .ActionText = "Received into Stock"
500           .UserName = UserName
510       End With
520       LogCourierInterface "SU3", MSG
530   Next

540   Exit Sub

SaveToBatchProducts_Error:

      Dim strES As String
      Dim intEL As Integer

550   intEL = Erl
560   strES = Err.Description
570   LogError "frmBatchProductEntry", "SaveToBatchProducts", intEL, strES
          
End Sub

Public Sub PrintBatchIdentifier(ByVal BatchNumber As String, _
                          ByVal Identifier As String, _
                          ByVal Product As String, _
                          ByVal Group As String)

10    On Error GoTo PrintBatchIdentifier_Error

20    If Not SetLabelPrinter() Then
30      Exit Sub
40    End If

50    Printer.Orientation = vbPRORPortrait
60    Printer.Font.Name = "3 of 9 Barcode"
70    Printer.CurrentX = 0
80    Printer.CurrentY = 100
90    Printer.CurrentX = 200
100   Printer.Font.Size = 18
110   Printer.Font.Bold = False
120   Printer.Print "*"; Identifier; "*"
130   Printer.Font.Bold = False
140   Printer.Font.Size = 10
150   Printer.Font.Name = "Courier New"
160   Printer.Print "Available for use"
170   Printer.Print Product
180   Printer.Print "Batch:"; BatchNumber
190   If Trim$(Group) <> "" Then
200     Printer.Print "Group:"; Group;
210   End If

220   Printer.EndDoc

230   Exit Sub

PrintBatchIdentifier_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "frmBatchProductEntry", "PrintBatchIdentifier", intEL, strES

End Sub

Private Sub dtExpiry_GotFocus()

10    cAuto.Caption = "Auto-Entry is OFF"
20    cAuto.ForeColor = vbRed
30    cAuto.Value = False

End Sub


Private Sub dtReceived_GotFocus()

10    cAuto.Caption = "Auto-Entry is OFF"
20    cAuto.ForeColor = vbRed
30    cAuto.Value = False

End Sub




Private Sub Form_Activate()

10    If tInput.Enabled Then
20      tInput.SetFocus
30    End If
        
End Sub

Private Sub Form_Load()

10    dtExpiry = Format(Now + 183, "dd/mmm/yyyy")
20    dtReceived = Format(Now, "dd/mmm/yyyy")
30    FillcmbProduct

40    With cGroup
50      .Clear
60      .AddItem "O"
70      .AddItem "A"
80      .AddItem "B"
90      .AddItem "AB"
100     .ListIndex = -1
110   End With

      '*****NOTE
          'this code might be dependent on many components so for any future
          'update in code try to keep this code on bottom most line of form load.
120       EnableControls False


      '**************************************

End Sub



Private Sub tInput_LostFocus()

10    On Error Resume Next
        
20    tInput = Trim$(UCase$(tInput))

30    lblReadError.Visible = False

      'Validate or Cancel
40    If tInput = CancelCode And CancelCode <> "" Then
50      cmdCancel_Click
60      Exit Sub
70    ElseIf tInput = ValidateCode And ValidateCode <> "" Then
80      cmdSave_Click
90      tInput = ""
100     tInput.SetFocus
110     Exit Sub
120   End If

130   Select Case Len(tInput)
        Case 0:
140       If Screen.ActiveControl.Tag = "cmdCancel" Then
150         cmdCancel_Click
160         Exit Sub
170       ElseIf Screen.ActiveControl.Tag = "cmdSave" Then
180         cmdSave_Click
190         tInput.SetFocus
200         Exit Sub
210       End If
220     Case 5: 'Group?
230       If Left$(tInput, 1) <> "D" Or Right$(tInput, 1) <> "B" Then
240         lblReadError.Visible = True
250       Else
260         Select Case Mid$(tInput, 2, 3)
              Case "550": cGroup = "O"
270           Case "660": cGroup = "A"
280           Case "770": cGroup = "B"
290           Case "880": cGroup = "AB"
300           Case Else: cGroup = ""
310                      lblReadError.Visible = True
320         End Select
330       End If
340       tInput = ""
350       tInput.SetFocus
          
360     Case 9: 'Product?
370       Select Case UCase$(tInput)
            Case "A0184903B": cmbProduct = "Octaplas"
380         Case "A0184803B": cmbProduct = "Uniplas"
390         Case Else: cmbProduct = ""
400                    lblReadError.Visible = True
410       End Select
420       tInput = ""
430       tInput.SetFocus
          
        
440     Case 10: 'Date? "a11022009c"
450       If UCase$(Left$(tInput, 1)) <> "A" Or UCase$(Right$(tInput, 1)) <> "C" Then
460         lblReadError.Visible = True
470       Else
480         dtExpiry = Convert62Date(Mid$(tInput, 2, 8), FORWARD)
490       End If
500       tInput = ""
510       tInput.SetFocus
          
520     Case 11: 'BatchNumber?
         'A6060139501
530       If Left$(tInput, 1) = "A" And Right$(tInput, 1) = "B" Then
540         txtBatchNumber = Mid$(tInput, 2, 9)
550       Else
560         txtBatchNumber = ""
570         lblReadError.Visible = True
580       End If
590       tInput = ""
600       tInput.SetFocus
        
610     Case Else:
620       lblReadError.Visible = True
630       tInput = ""
640       tInput.SetFocus
         
650   End Select

End Sub


Private Sub tVolume_GotFocus()

10    cAuto.Caption = "Auto-Entry is OFF"
20    cAuto.ForeColor = vbRed
30    cAuto.Value = False

End Sub


Private Sub txtBatchNumber_GotFocus()

10    cAuto.Caption = "Auto-Entry is OFF"
20    cAuto.ForeColor = vbRed
30    cAuto.Value = False

End Sub


Private Sub txtBatchNumber_LostFocus()

10    lblReadError.Visible = False

20    If Left$(txtBatchNumber, 1) = "A" And Right$(txtBatchNumber, 1) = "B" Then
30      txtBatchNumber = Mid$(txtBatchNumber, 2, 9)
40    End If

End Sub


Private Sub txtReceived_GotFocus()

10    cAuto.Caption = "Auto-Entry is OFF"
20    cAuto.ForeColor = vbRed
30    cAuto.Value = False

End Sub


Private Sub udReceived_GotFocus()

10    cAuto.Caption = "Auto-Entry is OFF"
20    cAuto.ForeColor = vbRed
30    cAuto.Value = False

End Sub


Private Sub udReceived_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    cAuto.Caption = "Auto-Entry is OFF"
20    cAuto.ForeColor = vbRed
30    cAuto.Value = False

End Sub


