VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "ComCt232.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "spin32.ocx"
Begin VB.Form frmNewStockBar 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Blood Stock Entry"
   ClientHeight    =   8160
   ClientLeft      =   525
   ClientTop       =   765
   ClientWidth     =   7200
   ControlBox      =   0   'False
   FillColor       =   &H0000C000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "7frmNewStockBar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8160
   ScaleWidth      =   7200
   Begin VB.Frame Frame3 
      Caption         =   "ISBT-128"
      Height          =   735
      Left            =   1710
      TabIndex        =   30
      Top             =   570
      Width           =   3405
      Begin VB.TextBox txtISBT128 
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   270
         Width           =   2715
      End
   End
   Begin VB.Frame fraOrder 
      Caption         =   "Order No."
      Height          =   975
      Left            =   180
      TabIndex        =   29
      Top             =   450
      Width           =   1485
      Begin VB.TextBox txtOrderNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         MaxLength       =   14
         TabIndex        =   0
         Tag             =   "tUnit"
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tag"
      Height          =   705
      Left            =   180
      TabIndex        =   26
      Top             =   6105
      Width           =   6885
      Begin VB.TextBox txtNotes 
         Height          =   285
         Left            =   420
         TabIndex        =   27
         Top             =   270
         Width           =   5985
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(Flag this message whenever this Unit is used)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   28
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdManualCodes 
      Caption         =   "Manual Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   630
      TabIndex        =   24
      Top             =   4530
      Width           =   5925
   End
   Begin VB.Frame Frame1 
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   150
      TabIndex        =   19
      Top             =   8550
      Width           =   6885
      Begin VB.TextBox tNotes 
         Height          =   285
         Left            =   420
         MaxLength       =   30
         TabIndex        =   20
         ToolTipText     =   "Free Text. Use for Local Supplier etc."
         Top             =   240
         Width           =   5955
      End
   End
   Begin VB.TextBox tInput 
      Height          =   285
      Left            =   7200
      TabIndex        =   17
      Top             =   1290
      Width           =   1875
   End
   Begin VB.Frame FrameCodes 
      Caption         =   "Codes"
      Enabled         =   0   'False
      Height          =   1125
      Left            =   180
      TabIndex        =   14
      Top             =   4845
      Width           =   6885
      Begin VB.TextBox tSpecial 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   420
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   630
         Width           =   5985
      End
      Begin VB.ComboBox cAntigen 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   420
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   270
         Width           =   5985
      End
   End
   Begin VB.Frame FrameSupplier 
      Caption         =   "Supplier"
      Enabled         =   0   'False
      Height          =   915
      Left            =   180
      TabIndex        =   13
      Top             =   3555
      Width           =   6885
      Begin VB.ComboBox cSupplier 
         Height          =   315
         Left            =   420
         TabIndex        =   5
         Top             =   360
         Width           =   5985
      End
   End
   Begin VB.Frame FrameProduct 
      Caption         =   "Product"
      Enabled         =   0   'False
      Height          =   795
      Left            =   180
      TabIndex        =   12
      Top             =   2565
      Width           =   6885
      Begin VB.ComboBox lstProduct 
         Height          =   315
         Left            =   420
         TabIndex        =   4
         Top             =   330
         Width           =   5985
      End
   End
   Begin VB.Frame FrameExpiry 
      Caption         =   "Expiry"
      Enabled         =   0   'False
      Height          =   975
      Left            =   2055
      TabIndex        =   11
      Top             =   1455
      Width           =   3195
      Begin ComCtl2.UpDown udExpiry 
         Height          =   525
         Left            =   1455
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   926
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtExpiry 
         Height          =   315
         Left            =   150
         MaxLength       =   16
         TabIndex        =   3
         Top             =   390
         Width           =   1245
      End
      Begin MSMask.MaskEdBox txtExpiryTime 
         Height          =   315
         Left            =   2145
         TabIndex        =   32
         Top             =   390
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin Spin.SpinButton sbHRSExpiry 
         Height          =   330
         Left            =   1875
         TabIndex        =   33
         Top             =   390
         Width           =   210
         _Version        =   65536
         _ExtentX        =   370
         _ExtentY        =   582
         _StockProps     =   73
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
      End
      Begin Spin.SpinButton sbMINExpiry 
         Height          =   330
         Left            =   2820
         TabIndex        =   34
         Top             =   390
         Width           =   210
         _Version        =   65536
         _ExtentX        =   370
         _ExtentY        =   582
         _StockProps     =   73
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
      End
   End
   Begin VB.Frame FrameGroup 
      Caption         =   "Group"
      Enabled         =   0   'False
      Height          =   975
      Left            =   180
      TabIndex        =   10
      Top             =   1455
      Width           =   1635
      Begin VB.ComboBox lstGroup 
         Height          =   330
         Left            =   180
         TabIndex        =   2
         Tag             =   "lstGroup"
         Top             =   390
         Width           =   1275
      End
   End
   Begin VB.Frame FrameUnit 
      Caption         =   "Unit Number"
      Enabled         =   0   'False
      Height          =   975
      Left            =   5160
      TabIndex        =   9
      Top             =   435
      Width           =   1695
      Begin VB.TextBox txtUnitNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         MaxLength       =   14
         TabIndex        =   31
         Tag             =   "tUnit"
         Top             =   345
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "&Validate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3960
      Picture         =   "7frmNewStockBar.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "bValidate"
      Top             =   7245
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5610
      Picture         =   "7frmNewStockBar.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "bCancel"
      Top             =   7275
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   180
      TabIndex        =   22
      Top             =   6915
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblGroupCheck 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Group Check is Disabled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   750
      TabIndex        =   25
      Top             =   7785
      Width           =   2895
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
      Left            =   150
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   6915
   End
   Begin VB.Label lAuto 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Auto-Entry is ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   750
      TabIndex        =   21
      Tag             =   "lAuto"
      Top             =   7275
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tInput ->"
      Height          =   255
      Left            =   6210
      TabIndex        =   18
      Top             =   1290
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Menu mOptions 
      Caption         =   "&Options"
      Begin VB.Menu mBlankGroup 
         Caption         =   "Blank &Group after entry"
      End
      Begin VB.Menu mBlankExpiry 
         Caption         =   "Blank &Expiry after entry"
      End
      Begin VB.Menu mBlankProduct 
         Caption         =   "Blank &Product after entry"
      End
      Begin VB.Menu mBlankSupplier 
         Caption         =   "Blank &Supplier after entry"
      End
   End
End
Attribute VB_Name = "frmNewStockBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean
Private AutoEntrySelection As Boolean

Private Sub EnableFrames(ByVal Enable As Boolean)

10        FrameUnit.Enabled = Enable
20        FrameGroup.Enabled = Enable
30        FrameExpiry.Enabled = Enable
40        FrameProduct.Enabled = Enable
50        FrameSupplier.Enabled = Enable
60        FrameCodes.Enabled = Enable

70        cmdManualCodes.Visible = False

End Sub

Private Sub GetSpecialCode()

          Dim n As Integer
10        ReDim Opt(0 To cAntigen.ListCount - 1)
          Dim f As Form

20        For n = 0 To cAntigen.ListCount - 1
30            Opt(n) = cAntigen.List(n)
40        Next

50        Set f = New fcdrDBox
60        With f
70            .Options = Opt
80            .Prompt = "Make Selection."
90            .Show 1
100           If TimedOut Then Unload Me: Exit Sub
110           If .ReturnValue <> "" Then
120               tSpecial = tSpecial & .ReturnValue & "."
130           End If
140       End With
150       Set f = Nothing

End Sub

Private Sub HighLightFrame(f As Frame)

10        With f
20            .Font.Bold = True
30            .Font.Size = 10
40            .BackColor = &HC0C0FF
50            .ForeColor = &HFF
60        End With

End Sub

Private Sub UnHighLightFrames()

          Dim f As Control

10        For Each f In Me.Controls
20            If TypeOf f Is Frame Then
30                With f
40                    .Font.Bold = False
50                    .Font.Size = 8
60                    .BackColor = &HC0C0C0
70                    .ForeColor = &H0
80                End With
90            End If
100       Next

End Sub


Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdValidate_Click()

          Dim s As String
          Dim Found As Boolean
          Dim ComponentCode As String
          Dim n As Integer
          Dim sql As String
          Dim tb As Recordset
          Dim EntryDateTime As String
          Dim Today As Date
          Dim datExpiry As Date
          Dim GroupChecked As Boolean
          Dim Generic As String
          Dim Ps As New Products
          Dim p As Product

          Dim strVisionGroupCheckTestProfile As String

10        On Error GoTo cmdValidate_Click_Error

20        If txtOrderNumber = "" Then
30            iMsg "Please enter order number first", vbInformation
40            If TimedOut Then Unload Me: Exit Sub
50            txtOrderNumber.SetFocus
60            Exit Sub
70        End If
80        If Not IsDate(txtExpiry & " " & txtExpiryTime) Or txtExpiry = "" Then
90            iMsg "Date of Expiry Entry Error", vbCritical
100           If TimedOut Then Unload Me: Exit Sub
110           txtExpiry = ""
120           If txtExpiry.Visible And txtExpiry.Enabled Then
130               txtExpiry.SetFocus
140           End If
150           Exit Sub
160       End If

170       datExpiry = Format(txtExpiry, "dd/mmm/yyyy")
180       Today = Now

190       If datExpiry - Today > 35 Or datExpiry - Today < 0 Then
200           Answer = iMsg("Is the Expiry Date correct?", vbQuestion + vbYesNo)
210           If TimedOut Then Unload Me: Exit Sub
220           If Answer = vbNo Then
230               txtExpiry = ""
240               txtExpiry.SetFocus
250               Exit Sub
260           End If
270       End If

280       If Trim$(txtISBT128) = "" Then
290           iMsg "Enter ISBT128", vbCritical
300           If TimedOut Then Unload Me: Exit Sub
310           txtISBT128.SetFocus
320           Exit Sub
330       Else
              'R000112902110 O
              'Validate IBST format
340           If Len(txtISBT128) <> 15 Then
350               iMsg "ISBT128 format incorrect", vbCritical
360               If TimedOut Then Unload Me: Exit Sub
370               txtISBT128.SetFocus
380               Exit Sub
390           End If

400           If Mid(txtISBT128, 14, 1) <> " " Then
410               iMsg "ISBT128 format incorrect", vbCritical
420               If TimedOut Then Unload Me: Exit Sub
430               txtISBT128.SetFocus
440               Exit Sub
450           End If

460           If Left(txtISBT128, 1) = "=" Then
470               iMsg "ISBT128 format incorrect", vbCritical
480               If TimedOut Then Unload Me: Exit Sub
490               txtISBT128.SetFocus
500               Exit Sub
510           End If
520       End If

530       If Trim$(lstproduct) = "" Then
540           iMsg "Select Component Type", vbCritical
550           If TimedOut Then Unload Me: Exit Sub
560           BlankDetails
570           tInput.SetFocus
580           Exit Sub
590       End If

600       ComponentCode = ProductBarCodeFor(lstproduct)
610       If ComponentCode = "???" Then
620           iMsg "Component not known!", vbCritical
630           If TimedOut Then Unload Me: Exit Sub
640           BlankDetails
650           tInput.SetFocus
660           Exit Sub
670       End If

680       If FrameUnit.Visible Then
690           If Trim$(txtUnitNumber) = "" Or Len(txtUnitNumber) < 6 Then
700               iMsg "Unit Number?", vbQuestion
710               If TimedOut Then Unload Me: Exit Sub
720               BlankDetails
730               tInput.SetFocus
740               Exit Sub
750           End If
760       End If

770       If Trim$(lstGroup) = "" Then
780           iMsg "Group not entered!", vbExclamation
790           If TimedOut Then Unload Me: Exit Sub
800           BlankDetails
810           tInput.SetFocus
820           Exit Sub
830       End If

840       Found = False
850       For n = 1 To 12
860           If lstGroup = Choose(n, "O Pos", "A Pos", "B Pos", "AB Pos", _
                                   "O Neg", "A Neg", "B Neg", "AB Neg", "O", "A", "B", "AB") Then
870               Found = True
880               Exit For
890           End If
900       Next
910       If Not Found Then
920           iMsg "Invalid Group/Rh!", vbExclamation
930           If TimedOut Then Unload Me: Exit Sub
940           BlankDetails
950           tInput.SetFocus
960           Exit Sub
970       End If

980       If Not IsDate(txtExpiry) Then
990           iMsg "Expiry date invalid.", vbCritical
1000          If TimedOut Then Unload Me: Exit Sub
1010          BlankDetails
1020          tInput.SetFocus
1030          Exit Sub
1040      End If

1050      If SupplierCodeFor(cSupplier) = "???" Then
1060          iMsg "Supplier?", vbQuestion
1070          If TimedOut Then Unload Me: Exit Sub
1080          BlankDetails
1090          tInput.SetFocus
1100          Exit Sub
1110      End If

1120      If InStr(tSpecial, "K+") = 0 And InStr(tSpecial, "K-") = 0 And _
             InStr(tSpecial, "K +") = 0 And InStr(tSpecial, "K -") = 0 And _
             InStr(tSpecial, "K P") = 0 And InStr(tSpecial, "K N") = 0 And _
             InStr(tSpecial, "Kell") = 0 Then
1130          Answer = iMsg("Is Unit Kell Status known?", vbQuestion + vbYesNo)
1140          If TimedOut Then Unload Me: Exit Sub
1150          If Answer = vbYes Then
1160              Exit Sub
1170          End If
1180      End If

1190      Ps.Load txtISBT128, ComponentCode
1200      If Ps.Count > 0 Then
1210          Set p = Ps.Item(1)
1220          s = "Product Number " & txtISBT128 & _
                " exists!" & vbCrLf & _
                  p.RecordDateTime & vbCrLf & _
                  gEVENTCODES(p.PackEvent).Text
1230          iMsg s, vbCritical
1240          If TimedOut Then Unload Me: Exit Sub
1250          BlankDetails
1260          tInput.SetFocus
1270          Exit Sub
1280      End If

1290      If DateDiff("d", Now, txtExpiry) < 0 Then
1300          Answer = iMsg("Unit expired. Do you wish to proceed?", vbExclamation + vbYesNo)
1310          If TimedOut Then Unload Me: Exit Sub
1320          If Answer = vbNo Then
1330              Exit Sub
1340          End If
1350      End If

1360      EntryDateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")

1370      If lblGroupCheck = "Group Check is Disabled" Then
1380          GroupChecked = True
1390      Else
1400          GroupChecked = False
1410      End If

1420      Set p = New Product
1430      p.PackNumber = UCase$(Replace(txtUnitNumber, "+", "X"))
1440      p.PackEvent = "C"
1450      p.Chart = ""
1460      p.PatName = ""
1470      p.UserName = UserCode
1480      p.RecordDateTime = EntryDateTime
1490      p.GroupRh = Group2Bar(lstGroup)
1500      p.Supplier = SupplierCodeFor(cSupplier.Text)
          'p.DateExpiry = Format(txtExpiry, "dd/MMM/yyyy hh:mm")
1510      p.DateExpiry = Format(txtExpiry & " " & txtExpiryTime, "dd/MMM/yyyy HH:mm")
1520      p.Screen = tSpecial
1530      p.SampleID = ""
1540      p.crt = 0
1550      p.cco = 0
1560      p.cen = 0
1570      p.crtr = 0
1580      p.ccor = 0
1590      p.cenr = 0
1600      p.Barcode = ProductBarCodeFor(lstproduct)
1610      p.Checked = GroupChecked
1620      p.Reason = ""
1630      p.OrderNumber = txtOrderNumber
1640      p.ISBT128 = txtISBT128
1650      p.Save

        'Check if unit details saved (Product/Latest)
        'If not then don't procees with Barcode, unit testing or Blood Track
1651      Ps.Load txtISBT128, ComponentCode
1652      If Ps.Count = 0 Then
1653        s = vbCrLf & "Product number " & txtISBT128 & vbCrLf & " not saved!"
1654        iMsg s, vbCritical
1655        If TimedOut Then Unload Me: Exit Sub
1656        BlankDetails
1657        tInput.SetFocus
1658        Exit Sub
1659      End If
 
1660      PrintBarCodesN txtISBT128, 1, "", "", "", lstGroup
1670      If Not GroupChecked Then
1675          strVisionGroupCheckTestProfile = GetOptionSetting("optVisionGroupCheckTestProfile", "UnitGroupCheck")
1680          sql = "Insert into BBOrderComms " & _
                    "(TestRequired, UnitNumber,SampleID) VALUES " & _
                    "('" & strVisionGroupCheckTestProfile & "', '" & Trim$(txtISBT128) & "', '')"
1690          CnxnBB(0).Execute sql
1700      End If

          Dim MSG As udtRS
1710      With MSG
1720          .UnitNumber = txtISBT128
1730          .ProductCode = ProductBarCodeFor(lstproduct)
1740          Generic = ProductGenericFor(.ProductCode)
1750          If Generic = "Platelets" Then
1760              .StorageLocation = strBTCourier_StorageLocation_RoomTempIssueFridge
1770          Else
1780              .StorageLocation = strBTCourier_StorageLocation_HemoSafeFridge
1790          End If
              '.UnitExpiryDate = Format(txtExpiry, "dd/mmm/yyyy")
1800          .UnitExpiryDate = Format(txtExpiry & " " & txtExpiryTime, "dd/MMM/yyyy HH:mm")
1810          .UnitGroup = lstGroup
              'No real Patient details when unit received
1820          .StockComment = ""
1830          .Chart = ""
1840          .PatientHealthServiceNumber = ""
1850          .ForeName = ""
1860          .SurName = ""
1870          .DoB = ""
1880          .Sex = ""
1890          .PatientGroup = ""
1900          .DeReservationDateTime = Format(txtExpiry & " " & txtExpiryTime, "dd-MMM-yyyy hh:mm:ss")
1910          .ActionText = "Received into Stock"
1920          .UserName = UserName
1930      End With
1940      LogCourierInterface "SU3", MSG

1950      If Trim$(txtNotes) = "" Then
1960          sql = "Delete from UnitNotes where " & _
                    "UnitNumber = '" & txtISBT128 & "' And DateExpiry = '" & Format(txtExpiry, "dd/mmm/yyyy") & "'"
1970          CnxnBB(0).Execute sql
1980      Else
1990          sql = "Select * from UnitNotes where " & _
                    "UnitNumber = '" & txtISBT128 & "' And DateExpiry = '" & Format(txtExpiry, "dd/mmm/yyyy") & "'"
2000          Set tb = New Recordset
2010          RecOpenClientBB 0, tb, sql
2020          If tb.EOF Then
2030              tb.AddNew
2040          End If
2050          tb!Notes = txtNotes
2060          tb!UnitNumber = txtISBT128
2070          tb!DateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
2080          tb!Technician = UserName
              'tb!DateExpiry = Format(txtExpiry, "dd/mmm/yyyy")
2090          tb!DateExpiry = Format(txtExpiry & " " & txtExpiryTime, "dd/MMM/yyyy HH:mm")
2100          tb.Update
2110      End If

2120      If tInput.Visible Then
2130          tInput.SetFocus
2140      End If

2150      BlankDetails

          'Turn Auto-Entry ON (after manual entry)
2160      If lauto = "Auto-Entry is OFF" Then
2170          AutoEntryOnOff True
2180          AutoEntrySelection = True

2190          tInput = ""
2200          tInput.SetFocus
2210      End If

2220      Exit Sub

cmdValidate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

2230      intEL = Erl
2240      strES = Err.Description
2250      LogError "frmNewStockBar", "cmdValidate_Click", intEL, strES, sql

End Sub
Private Sub BlankDetails()

10        If mBlankExpiry.Checked Then txtExpiry = ""
20        If mBlankGroup.Checked Then lstGroup = ""
30        cAntigen = ""
40        If mBlankProduct.Checked Then lstproduct = ""
50        If mBlankSupplier.Checked Then cSupplier = ""
60        txtUnitNumber = ""
70        tSpecial = ""
80        tNotes = ""
90        txtNotes = ""
100       txtISBT128 = ""

End Sub

Private Sub cantigen_Click()

10        If Trim$(cAntigen) <> "" Then
20            tSpecial = tSpecial & cAntigen & ". "
30            cAntigen = ""
40        End If

End Sub

Private Sub cantigen_GotFocus()

10        UnHighLightFrames
20        HighLightFrame FrameCodes

End Sub


Private Sub cAntigen_LostFocus()

          Dim s As String

10        On Error Resume Next

20        If Trim$(cAntigen) <> "" Then
30            s = AntigenDescription(cAntigen)
40            If s <> "" Then
50                tSpecial = tSpecial & s & ". "
60            Else
70                tSpecial = tSpecial & cAntigen & ". "
80            End If
90        End If

100       cAntigen = ""

110       UnHighLightFrames

End Sub

Private Sub cmdManualCodes_Click()

10        GetSpecialCode

End Sub

Private Sub cSupplier_GotFocus()

10        UnHighLightFrames
20        HighLightFrame FrameSupplier

End Sub


Private Sub cSupplier_LostFocus()


10        If Left$(cSupplier, 2) = "A0" Then
20            cSupplier = Mid$(cSupplier, 3, 7)
30            cSupplier = SupplierNameFor(cSupplier)
40        End If

50        UnHighLightFrames

End Sub


Private Sub Form_Activate()

      '10    Me.Width = 7290

10        If Activated Then Exit Sub

20        If tInput.Enabled Then
30            tInput.SetFocus
40        End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

10        If KeyAscii = 5 Then KeyAscii = 55

End Sub

Private Sub Form_Load()

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo Form_Load_Error

20        LoadSpecialTestingSettings

30        pubStrStockSupplierName_IRL = GetOptionSetting("optStockSupplierName_IRL", "IBTS Dublin")

40        AutoEntrySelection = True

50        mBlankGroup.Checked = GetSetting("NetAcquire", "Transfusion", "BlankGroup", True)
60        mBlankExpiry.Checked = GetSetting("NetAcquire", "Transfusion", "BlankExpiry", True)
70        mBlankProduct.Checked = GetSetting("NetAcquire", "Transfusion", "BlankProduct", True)
80        mBlankSupplier.Checked = GetSetting("NetAcquire", "Transfusion", "BlankSupplier", True)

90        FrameUnit.Visible = GetOptionSetting("optISBT128_Save7DigitUnitNumbers", "1")



100       For n = 0 To 13
110           lstGroup.AddItem Choose(n + 1, " ", "O Pos", "A Pos", _
                                      "B Pos", "AB Pos", _
                                      "O Neg", "A Neg", _
                                      "B Neg", "AB Neg", "O", "A", "B", "AB", "??"), n
120       Next

130       sql = "Select Wording from ProductList order by ListOrder"
140       Set tb = New Recordset
150       RecOpenServerBB 0, tb, sql
160       Do While Not tb.EOF
170           lstproduct.AddItem tb!Wording
180           tb.MoveNext
190       Loop

200       sql = "Select * from Supplier " & _
                "Order by ListOrder"
210       Set tb = New Recordset
220       RecOpenServerBB 0, tb, sql
230       Do While Not tb.EOF
240           cSupplier.AddItem tb!Supplier & ""
250           tb.MoveNext
260       Loop

270       sql = "Select * from Lists where " & _
                "ListType = 'PC' " & _
                "Order by ListOrder"
280       Set tb = New Recordset
290       RecOpenServerBB 0, tb, sql
300       Do While Not tb.EOF
310           cAntigen.AddItem tb!Text & ""
320           tb.MoveNext
330       Loop
340       txtExpiryTime = "23:59"

350       EnableFrames False

360       lblGroupCheck.Visible = True

370       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

380       intEL = Erl
390       strES = Err.Description
400       LogError "frmNewStockBar", "Form_Load", intEL, strES, sql

End Sub

Private Sub FrameCodes_Click()

10        UnHighLightFrames
20        HighLightFrame FrameCodes
30        cAntigen.SetFocus

End Sub

Private Sub FrameExpiry_Click()

10        UnHighLightFrames
20        HighLightFrame FrameExpiry
30        txtExpiry.SetFocus

End Sub

Private Sub FrameGroup_Click()

10        UnHighLightFrames
20        HighLightFrame FrameGroup
30        lstGroup.SetFocus

End Sub

Private Sub FrameProduct_Click()

10        UnHighLightFrames

20        HighLightFrame FrameProduct
30        lstproduct.SetFocus

End Sub

Private Sub FrameSupplier_Click()

10        UnHighLightFrames
20        HighLightFrame FrameSupplier
30        cSupplier.SetFocus

End Sub

Private Sub FrameUnit_Click()

10        UnHighLightFrames
20        HighLightFrame FrameUnit
30        txtUnitNumber.SetFocus

End Sub

Private Sub fraOrder_Click()
10        UnHighLightFrames
20        HighLightFrame fraOrder
30        txtOrderNumber.SetFocus
End Sub

Private Sub Label1_Click()

10        tInput.SetFocus

End Sub

Private Sub lAuto_Click()
          Dim s As String

10        If lauto = "Auto-Entry is ON" Then
20            s = "Turn Auto-Entry OFF?"
30            Answer = iMsg(s, vbYesNo + vbQuestion)
40            If TimedOut Then Unload Me: Exit Sub
50            If Answer = vbYes Then
60                AutoEntrySelection = False
70                AutoEntryOnOff False
80                txtOrderNumber.SetFocus
90            End If
100       Else
110           s = "Turn Auto-Entry ON?"
120           Answer = iMsg(s, vbYesNo + vbQuestion)
130           If TimedOut Then Unload Me: Exit Sub
140           If Answer = vbYes Then

150               AutoEntryOnOff True
160               AutoEntrySelection = True

170               tInput = ""
180               tInput.SetFocus
190           End If
200       End If

End Sub

Private Sub AutoEntryOnOff(Value As Boolean)
10        lauto = IIf(Value, "Auto-Entry is ON", "Auto-Entry is OFF")
20        lauto.ForeColor = IIf(Value, vbGreen, vbRed)
30        EnableFrames IIf(Value, False, True)
End Sub


Private Sub lblGroupCheck_Click()

10        If lblGroupCheck = "Group Check is Enabled" Then
20            lblGroupCheck = "Group Check is Disabled"
30            lblGroupCheck.ForeColor = vbRed
40        Else
50            lblGroupCheck = "Group Check is Enabled"
60            lblGroupCheck.ForeColor = vbGreen
70        End If

End Sub

Private Sub lstgroup_GotFocus()

10        UnHighLightFrames
20        HighLightFrame FrameGroup

End Sub


Private Sub lstgroup_LostFocus()

          Dim Temp As String

10        Temp = Mid$(lstGroup, 2, 2)
20        Temp = Bar2Group(Temp)
30        If Temp = "" Then
40            Temp = Group2Bar(lstGroup)
50            If Temp = "" Then
60                lstGroup = ""
70            End If
80        Else
90            lstGroup = Temp
100       End If

110       UnHighLightFrames

End Sub

Private Sub lstproduct_GotFocus()

10        UnHighLightFrames

20        HighLightFrame FrameProduct

End Sub



Private Sub lstproduct_LostFocus()

'          Dim Temp As String
'
'10        Temp = lstProduct
'
'20        If Len(Temp) = 9 Then
'30            lstProduct = Mid$(Temp, 3, 5)
'40            lstProduct = ProductWordingFor(lstProduct)
'50        End If

End Sub

Private Sub mBlankExpiry_Click()

10        mBlankExpiry.Checked = Not mBlankExpiry.Checked
20        SaveSetting "NetAcquire", "Transfusion", "BlankExpiry", CStr(mBlankExpiry.Checked)

End Sub

Private Sub mBlankGroup_Click()

10        mBlankGroup.Checked = Not mBlankGroup.Checked
20        SaveSetting "NetAcquire", "Transfusion", "BlankGroup", CStr(mBlankGroup.Checked)

End Sub

Private Sub mBlankProduct_Click()

10        mBlankProduct.Checked = Not mBlankProduct.Checked
20        SaveSetting "NetAcquire", "Transfusion", "BlankProduct", CStr(mBlankProduct.Checked)

End Sub

Private Sub mBlankSupplier_Click()

10        mBlankSupplier.Checked = Not mBlankSupplier.Checked
20        SaveSetting "NetAcquire", "Transfusion", "BlankSupplier", CStr(mBlankSupplier.Checked)

End Sub
Private Sub sbMINExpiry_SpinUp()
          Dim intRight As Integer
          Dim strL As String
          Dim strR As String

10        On Error GoTo sbMINExpiry_SpinUp_Error

20        strL = Left(txtExpiryTime, 2)
30        strR = Right(txtExpiryTime, 2)

40        If InStr(strL, "_") Then Exit Sub
50        If InStr(strR, "_") Then Exit Sub

60        intRight = Int(strR)

70        If intRight < 59 Then
80            intRight = intRight + 1
90        Else
100           intRight = 0
110       End If

120       strR = Format(intRight, "00")
130       txtExpiryTime = strL & ":" & strR

140       Exit Sub

sbMINExpiry_SpinUp_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "fNewStockBar", "sbMINExpiry_SpinUp", intEL, strES

End Sub
Private Sub sbMINExpiry_SpinDown()
          Dim intRight As Integer
          Dim strL As String
          Dim strR As String

10        strL = Left(txtExpiryTime, 2)
20        strR = Right(txtExpiryTime, 2)

30        If InStr(strL, "_") Then Exit Sub
40        If InStr(strR, "_") Then Exit Sub

50        intRight = Int(strR)

60        If intRight > 0 Then
70            intRight = intRight - 1
80        Else
90            intRight = 59
100       End If

110       strR = Format(intRight, "00")
120       txtExpiryTime = strL & ":" & strR
End Sub
Private Sub sbHRSExpiry_SpinUp()
          Dim intLeft As Integer
          Dim strL As String
          Dim strR As String

10        On Error GoTo sbHRSExpiry_SpinUp_Error

20        strL = Left(txtExpiryTime, 2)
30        strR = Right(txtExpiryTime, 2)

40        If InStr(strL, "_") Then Exit Sub
50        If InStr(strR, "_") Then Exit Sub

60        intLeft = Int(strL)

70        If intLeft < 23 Then
80            intLeft = intLeft + 1
90        Else
100           intLeft = 0
110       End If

120       strL = Format(intLeft, "00")
130       txtExpiryTime = strL & ":" & strR

140       Exit Sub

sbHRSExpiry_SpinUp_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "fNewStockBar", "sbHRSExpiry_SpinUp", intEL, strES

End Sub
Private Sub sbHRSExpiry_SpinDown()
          Dim intLeft As Integer
          Dim strL As String
          Dim strR As String

10        On Error GoTo sbHRSExpiry_SpinDown_Error

20        strL = Left(txtExpiryTime, 2)
30        strR = Right(txtExpiryTime, 2)

40        If InStr(strL, "_") Then Exit Sub
50        If InStr(strR, "_") Then Exit Sub

60        intLeft = Int(strL)

70        If intLeft > 0 Then
80            intLeft = intLeft - 1
90        Else
100           intLeft = 23
110       End If

120       strL = Format(intLeft, "00")
130       txtExpiryTime = strL & ":" & strR

140       Exit Sub

sbHRSExpiry_SpinDown_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "fNewStockBar", "sbHRSExpiry_SpinDown", intEL, strES

End Sub



Private Sub txtExpiry_GotFocus()

10        UnHighLightFrames
20        HighLightFrame FrameExpiry
30        txtExpiry.SelStart = 0
40        txtExpiry.SelLength = Len(txtExpiry)

End Sub

Private Sub txtexpiry_LostFocus()

      'txtExpiry = Convert62Date(txtExpiry, FORWARD)

10        UnHighLightFrames

End Sub

Private Sub ParseIBTS()
      ''===============================
          Dim code As String
          Dim s As String
10        Select Case Left(tInput, 2)
          Case "&>"
              'Expiry Date
20            code = Mid(tInput, 4, 9)
30            txtExpiry = DateAdd("d", CDbl(Mid(code, 3, 3)) - 1, "01/01/" & Mid(code, 1, 2))
40            txtExpiryTime = (Mid(code, 6, 2) & ":" & Mid(code, 8, 2))
50            tInput = ""
60            tInput.SetFocus
70        Case "=<"
              'Product
80            code = Mid(tInput, 3, Len(tInput))
90            lstproduct = ProductWordingFor(code)
100           tInput = ""
110           tInput.SetFocus
120       Case "=%"
              'Blood Group
130           code = Mid(tInput, 3, 2)
140           lstGroup = Bar2Group(code)
150           tInput = ""
160           tInput.SetFocus
170       Case "&{"
180           tSpecial = PlateletHLAandPlateletSpecificAntigens(tInput)
190           tInput = ""
200           tInput.SetFocus
210       Case "=\", "=#"
              'Special Testing: Red Blood Cell Antigens -- General
220           tSpecial = RBCAntigensGeneral(tInput)
230           tInput = ""
240           tInput.SetFocus
250       Case Else
260           Select Case Left(tInput, 1)
              Case "="
270               If Len(tInput) = 16 Then
                      'Donation Identifier Number
280                   code = Mid(tInput, 2, 13)
290                   s = ISOmod37_2(Mid$(tInput, 2, 13))
300                   txtISBT128 = code & " " & s

310                   If Mid(Trim$(tInput), 2, 5) = "R0001" Or Mid(Trim$(tInput), 2, 5) = "X0002" Then
320                       cSupplier = pubStrStockSupplierName_IRL
330                   End If

340                   tInput = ""
350                   tInput.SetFocus
360               End If

370           Case "&"
                  'Unstandarised barcode not from ICCBBA
380               Select Case Mid(tInput, 2, 1)
                      ' Unit Number
                  Case "a", "A"
390                   code = Mid(tInput, 3, Len(tInput))
400                   txtUnitNumber = code
410                   tInput = ""
420                   tInput.SetFocus
430               End Select
440           End Select
450       End Select
End Sub

Private Sub tInput_LostFocus()

          Dim s As String
          Dim tb As Recordset
          Dim sql As String
          Dim Check As String
          Dim Ps As New Products
          Dim p As Product

10        On Error GoTo tInput_LostFocus_Error


20        lblReadError.Visible = False

30        tInput = Replace(tInput, "'", "")

40        tInput = UCase$(Trim$(tInput))

50        If tInput = "" Then Exit Sub

60        If Screen.ActiveControl.Tag = "cmdCancel" Then
70            cmdCancel_Click
80            Exit Sub
90        End If
100       If Screen.ActiveControl.Tag = "lAuto" Then
110           lAuto_Click
120           Exit Sub
130       End If

140       If Left(tInput, 1) = "=" Or Left(tInput, 1) = "&" Then
              'Its IBTS code
150           ParseIBTS
160           Exit Sub
170       End If
          'Validate or Cancel
180       If tInput = ValidateCode Then
190           cmdValidate_Click
200           tInput = ""
210           tInput.SetFocus
220           Exit Sub
230       ElseIf tInput = CancelCode Then
240           Unload Me
250           Exit Sub
260       Else
270           sql = "Select * from Lists where " & _
                    "Code = '" & AddTicks(tInput) & "'"
280           Set tb = New Recordset
290           RecOpenServerBB 0, tb, sql
300           If Not tb.EOF Then
310               tSpecial = tSpecial & tb!Text & ". "
320               tInput = ""
330               tInput.SetFocus
340               Exit Sub
350           End If
360       End If

370       Select Case Len(tInput)

          Case 5:    'Group?
380           If Left$(tInput, 1) <> "D" Or Right$(tInput, 1) <> "B" Then
390               lblReadError.Visible = True
400           Else
410               lstGroup = Bar2Group(Mid$(tInput, 2, 2))
420           End If
430           tInput = ""
440           tInput.SetFocus

450       Case 6:    'Antigen?
460           If Left$(tInput, 1) <> "A" Then
470               lblReadError.Visible = True
480           Else
490               s = AntigenDescription(Mid$(tInput, 2, 4))
500               If s <> "" Then
510                   tSpecial = tSpecial & s & ". "
520               Else
530                   iMsg "Antigen not known!" & vbCrLf & _
                           "Go to ""Lists"" - ""Antigens""" & vbCrLf & _
                           "to enter a new Antigen Code", vbExclamation
540                   If TimedOut Then Unload Me: Exit Sub
550               End If
560           End If
570           tInput = ""
580           tInput.SetFocus

590       Case 8:    'Date?
600           tInput = CheckJulian(tInput)
610           If IsDate(tInput) Then
620               txtExpiry = tInput
630               tInput = ""
640               tInput.SetFocus
650               Exit Sub
660           End If
670           If Left$(tInput, 1) <> "A" Or Right$(tInput, 1) <> "C" Then
680               lblReadError.Visible = True
690           Else
700               txtExpiry = Convert62Date(Mid$(tInput, 2, 6), FORWARD)
710           End If
720           tInput = ""
730           tInput.SetFocus

740       Case 9:    'Pack ID or Product?
750           tInput = UCase$(Replace(tInput, "+", "X"))
760           If Left$(tInput, 1) = "D" And Right$(tInput, 1) = "D" Then    'pack
770               txtUnitNumber = Mid$(tInput, 2, 7)
780               Check = ChkDig(Left$(txtUnitNumber, 6))
790               If Check <> Right$(txtUnitNumber, 1) Then
800                   iMsg "Check Digit incorrect!", vbCritical
810                   If TimedOut Then Unload Me: Exit Sub
820                   txtUnitNumber = ""
830               Else

840                   Ps.LoadLatestByUnitNumber txtUnitNumber
850                   If Ps.Count > 0 Then
860                       Set p = Ps(1)
870                       s = "Unit Number already used!" & vbCrLf & _
                              "Expiry " & Format$(p.DateExpiry, "dd/MM/yyyy HH:mm") & vbCrLf & _
                              "Product " & ProductWordingFor(p.Barcode) & vbCrLf & _
                              "ISBT128 : " & p.ISBT128 & vbCrLf & _
                              "Continue?"
880                       Answer = iMsg(s, vbYesNo + vbQuestion, "Duplicate Number", vbRed)
890                       If TimedOut Then Unload Me: Exit Sub
900                       If Answer = vbNo Then
910                           txtUnitNumber = ""
920                       End If
930                   End If
940               End If

950           ElseIf Left$(tInput, 1) = "A" And Right$(tInput, 1) = "B" Then    'Product
960               lstproduct = ProductWordingFor(Mid$(tInput, 3, 5))
970           Else
980               lblReadError.Visible = True
990           End If
1000          tInput = ""
1010          tInput.SetFocus

1020      Case 10:    'Julian Date?
1030          tInput = CheckJulian(tInput)
1040          If Not IsDate(tInput) Then
1050              lblReadError.Visible = True
1060          Else
1070              txtExpiry = tInput
1080          End If
1090          tInput = ""
1100          tInput.SetFocus

1110      Case 11:    'Supplier?
1120          If Left$(tInput, 1) <> "A" Or Right$(tInput, 1) <> "B" Then
1130              lblReadError.Visible = True
1140          Else
1150              cSupplier = SupplierNameFor(Mid$(tInput, 3, 7))
1160          End If
1170          tInput = ""
1180          tInput.SetFocus

1190      Case 16:    'ISBT128
1200          If Left$(tInput, 1) <> "=" Then
1210              lblReadError.Visible = True
1220          Else
1230              s = ISOmod37_2(Mid$(tInput, 2, 13))
1240              txtISBT128 = Mid$(tInput, 2, 13) & " " & s
1250          End If
1260          tInput = ""
1270          tInput.SetFocus

1280      Case Else:
1290          lblReadError.Visible = True
1300          tInput = ""
1310          tInput.SetFocus

1320      End Select


1330      Exit Sub


          '===============================
1340      Exit Sub

tInput_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

1350      intEL = Erl
1360      strES = Err.Description
1370      LogError "frmNewStockBar", "tInput_LostFocus", intEL, strES, sql

End Sub



Private Sub txtOrderNumber_GotFocus()
10        UnHighLightFrames
20        HighLightFrame fraOrder
30        txtOrderNumber.SelStart = 0
40        txtOrderNumber.SelLength = Len(txtOrderNumber)

50        If AutoEntrySelection = True Then AutoEntryOnOff False

End Sub

Private Sub txtOrderNumber_LostFocus()
10        If AutoEntrySelection = True Then
20            AutoEntryOnOff True
30            tInput = ""
40            tInput.SetFocus
50        End If
End Sub


Private Sub txtUnitNumber_GotFocus()

10        UnHighLightFrames
20        HighLightFrame FrameUnit
30        txtUnitNumber.SelStart = 0
40        txtUnitNumber.SelLength = Len(txtUnitNumber)

End Sub

Private Sub txtUnitNumber_LostFocus()

          Dim s As String
          Dim Temp As String
          Dim Check As String
          Dim sql As String
          Dim Ps As New Products
          Dim p As Product

10        On Error GoTo txtUnitNumber_LostFocus_Error

20        UnHighLightFrames

30        If Len(Trim$(txtUnitNumber)) = 0 Then
40            Exit Sub
50        End If

60        txtUnitNumber = Trim$(UCase$(Replace(txtUnitNumber, "+", "X")))

70        If Len(txtUnitNumber) > 7 Then
80            txtUnitNumber = Mid$(txtUnitNumber, 2, 7)
90            Check = ChkDig(Left$(txtUnitNumber, 6))
100           If Check <> Right$(txtUnitNumber, 1) Then
110               iMsg "Check Digit incorrect!", vbCritical
120               If TimedOut Then Unload Me: Exit Sub
130               txtUnitNumber = ""
140               Exit Sub
150           End If
160       ElseIf Len(txtUnitNumber) = 7 Then
170           Check = ChkDig(Left$(txtUnitNumber, 6))
180           If Check <> Right$(txtUnitNumber, 1) Then
190               iMsg "Check Digit incorrect!", vbCritical
200               If TimedOut Then Unload Me: Exit Sub
210               txtUnitNumber = ""
220               Exit Sub
230           End If
240       End If

250       Temp = Convert62Date(txtUnitNumber, DONTCARE)
260       If IsDate(Temp) Then
270           s = "Unit number could represent date " & _
                  Format(Temp, "dd/mm/yyyy") & " " & _
                  "Is the Unit Number Correct?"
280           Answer = iMsg(s, vbYesNo + vbQuestion)
290           If TimedOut Then Unload Me: Exit Sub
300           If Answer = vbNo Then
310               txtUnitNumber = ""
320               txtUnitNumber.SetFocus
330               Exit Sub
340           End If
350       End If

360       If Len(Trim$(txtUnitNumber)) = 0 Then
370           Exit Sub
380       End If

390       Ps.LoadLatestByUnitNumber txtUnitNumber
400       If Ps.Count > 0 Then
410           Set p = Ps(1)
420           s = "Unit Number already used!" & vbCrLf & _
                  "Expiry " & Format$(p.DateExpiry, "dd/MM/yyyy HH:mm") & vbCrLf & _
                  "Product " & ProductWordingFor(p.Barcode) & vbCrLf & _
                  "ISBT128 : " & p.ISBT128 & vbCrLf & _
                  "Codabar : " & p.PackNumber & vbCrLf & _
                  "Continue?"
430           Answer = iMsg(s, vbYesNo + vbQuestion, "Duplicate Number", vbRed)
440           If TimedOut Then Unload Me: Exit Sub
450           If Answer = vbNo Then
460               txtUnitNumber = ""
470           End If
480       End If

490       Exit Sub

txtUnitNumber_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

500       intEL = Erl
510       strES = Err.Description
520       LogError "frmNewStockBar", "txtUnitNumber_LostFocus", intEL, strES, sql

End Sub

Private Sub udExpiry_DownClick()

10        If IsDate(txtExpiry) Then
20            txtExpiry = Format(DateAdd("d", -1, txtExpiry), "dd/mm/yyyy")
30        Else
40            txtExpiry = Format(Now, "dd/mm/yyyy")
50        End If

End Sub


Private Sub udExpiry_UpClick()

10        If IsDate(txtExpiry) Then
20            txtExpiry = Format(DateAdd("d", 1, txtExpiry), "dd/mm/yyyy")
30        Else
40            txtExpiry = Format(Now, "dd/mm/yyyy")
50        End If

End Sub


Private Sub LoadSpecialTestingSettings()

      'RBCRT009Interpretation
      'https://www.iccbba.org/uploads/e3/3c/e33c12d0ba56e7b63f87e9a9c5d2f48f/ST-001-ISBT-128-Standard-Technical-Specification-v5.2.0.pdf
      'Data structure 012: Special Testing: Red Blood Cell Antigens -- General, Positions 1 through 9 [RT009]
      'Red Blood Cell Antigens - Position 1 - Values 0 to 9
10    On Error GoTo LoadSpecialTestingSettings_Error

20        pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val0 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val0", "C+c-E+e-")
30        pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val1", "C+c+E+e-")
40        pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val2", "C-c+E+e-")
50        pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val3 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val3", "C+c-E+e+")
60        pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val4 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val4", "C+c+E+e+")
70        pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val5 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val5", "C-c+E+e+")
80        pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val6 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val6", "C+c-E-e+")
90        pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val7 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val7", "C+c+E-e+")
100       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val8 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val8", "C-c+E-e+")
110       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val9 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val9", "")
          'Red Blood Cell Antigens - Position 2 - Antigens
120       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos2_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos2_Antigen1", "K")
130       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos2_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos2_Antigen2", "k")
          'Red Blood Cell Antigens - Position 3 - Antigens
140       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos3_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos3_Antigen1", "Cw")
150       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos3_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos3_Antigen2", "Mia")
          'Red Blood Cell Antigens - Position 4 - Antigens
160       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos4_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos4_Antigen1", "M")
170       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos4_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos4_Antigen2", "N")
          'Red Blood Cell Antigens - Position 5 - Antigens
180       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos5_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos5_Antigen1", "S")
190       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos5_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos5_Antigen2", "s")
          'Red Blood Cell Antigens - Position 6 - Antigens
200       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos6_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos6_Antigen1", "U")
210       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos6_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos6_Antigen2", "P1")
          'Red Blood Cell Antigens - Position 7 - Antigens
220       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos7_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos7_Antigen1", "Lua")
230       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos7_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos7_Antigen2", "Kpa")
          'Red Blood Cell Antigens - Position 8 - Antigens
240       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos8_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos8_Antigen1", "Lea")
250       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos8_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos8_Antigen2", "Leb")
          'Red Blood Cell Antigens - Position 9 - Antigens
260       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos9_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos9_Antigen1", "Fya")
270       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos9_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos9_Antigen2", "Fyb")
          'Red Blood Cell Antigens - Position 10 - Antigens
280       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos10_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos10_Antigen1", "Jka")
290       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos10_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos10_Antigen2", "Jkb")
          'Red Blood Cell Antigens - Position 11 - Antigens
300       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos11_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos11_Antigen1", "Doa")
310       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos11_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos11_Antigen2", "Dob")
          'Red Blood Cell Antigens - Position 12 - Antigens
320       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos12_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos12_Antigen1", "Ina")
330       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos12_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos12_Antigen2", "Cob")
          'Red Blood Cell Antigens - Position 13 - Antigens
340       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos13_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos13_Antigen1", "Dia")
350       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos13_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos13_Antigen2", "VS/V")
          'Red Blood Cell Antigens - Position 14 - Antigens
360       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos14_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos14_Antigen1", "Jsa")
370       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos14_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos14_Antigen2", "C")
          'Red Blood Cell Antigens - Position 15 - Antigens
380       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos15_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos15_Antigen1", "c")
390       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos15_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos15_Antigen2", "E")
          'Red Blood Cell Antigens - Position 16 - Antigens
400       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos16_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos16_Antigen1", "e")
410       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos16_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Pos16_Antigen2", "CMV")

          'Red Blood Cell Antigens - Antigen Values
420       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value0_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value0_Antigen1", "")
430       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value0_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value0_Antigen2", "")

440       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value1_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value1_Antigen1", "")
450       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value1_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value1_Antigen2", "-")

460       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value2_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value2_Antigen1", "")
470       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value2_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value2_Antigen2", "+")

480       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value3_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value3_Antigen1", "-")
490       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value3_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value3_Antigen2", "")

500       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value4_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value4_Antigen1", "-")
510       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value4_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value4_Antigen2", "-")

520       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value5_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value5_Antigen1", "-")
530       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value5_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value5_Antigen2", "+")

540       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value6_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value6_Antigen1", "+")
550       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value6_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value6_Antigen2", "")

560       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value7_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value7_Antigen1", "+")
570       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value7_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value7_Antigen2", "-")

580       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value8_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value8_Antigen1", "+")
590       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value8_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value8_Antigen2", "+")

600       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value9_Antigen1 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value8_Antigen9", "")
610       pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value9_Antigen2 = GetOptionSetting("DS12SpecialTestingRedBloodCellAntigensGeneral_Value8_Antigen9", "")



          'https://www.iccbba.org/uploads/e3/3c/e33c12d0ba56e7b63f87e9a9c5d2f48f/ST-001-ISBT-128-Standard-Technical-Specification-v5.2.0.pdf
          'PlateletHLAandPlateletSpecificAntigens
          'Special Testing: Platelets HLA and Platelet Specific Antigens [Data Structure 014]
          'Data Structure 014:Special Testing: Platelets -Specific Antigens , Position 9 through 16 [RT014]
620       pubStrRT014Interpretation_Pos9r1 = GetOptionSetting("RT014Interpretation_Pos9R1", "HPA-1a")
630       pubStrRT014Interpretation_Pos9r2 = GetOptionSetting("RT014Interpretation_Pos9R2", "HPA-1b")

640       pubStrRT014Interpretation_Pos10r1 = GetOptionSetting("RT014Interpretation_Pos10R1", "HPA-2a")
650       pubStrRT014Interpretation_Pos10r2 = GetOptionSetting("RT014Interpretation_Pos10R2", "HPA-2b")

660       pubStrRT014Interpretation_Pos11r1 = GetOptionSetting("RT014Interpretation_Pos11R1", "HPA-3a")
670       pubStrRT014Interpretation_Pos11r2 = GetOptionSetting("RT014Interpretation_Pos11R2", "HPA-3b")

680       pubStrRT014Interpretation_Pos12r1 = GetOptionSetting("RT014Interpretation_Pos12R1", "HPA-4a")
690       pubStrRT014Interpretation_Pos12r2 = GetOptionSetting("RT014Interpretation_Pos12R2", "HPA-4b")

700       pubStrRT014Interpretation_Pos13r1 = GetOptionSetting("RT014Interpretation_Pos13R1", "HPA-5a")
710       pubStrRT014Interpretation_Pos13r2 = GetOptionSetting("RT014Interpretation_Pos13R2", "HPA-5b")

720       pubStrRT014Interpretation_Pos14r1 = GetOptionSetting("RT014Interpretation_Pos14R1", "HPA-15a")
730       pubStrRT014Interpretation_Pos14r2 = GetOptionSetting("RT014Interpretation_Pos14R2", "HPA-6bw")

740       pubStrRT014Interpretation_Pos15r1 = GetOptionSetting("RT014Interpretation_Pos15R1", "HPA-15b")
750       pubStrRT014Interpretation_Pos15r2 = GetOptionSetting("RT014Interpretation_Pos15R2", "HPA-7bw")

760       pubStrRT014Interpretation_Pos16r1 = GetOptionSetting("RT014Interpretation_Pos16R1", "IgA")
770       pubStrRT014Interpretation_Pos16r2 = GetOptionSetting("RT014Interpretation_Pos16R2", "CMV")

780       pubStrRT014Interpretation_Value1r1 = GetOptionSetting("RT014Interpretation_Value1r1", "")
790       pubStrRT014Interpretation_Value1r2 = GetOptionSetting("RT014Interpretation_Value1r2", "Neg")
800       pubStrRT014Interpretation_Value2r1 = GetOptionSetting("RT014Interpretation_Value2r1", "")
810       pubStrRT014Interpretation_Value2r2 = GetOptionSetting("RT014Interpretation_Value2r2", "Pos")
820       pubStrRT014Interpretation_Value3r1 = GetOptionSetting("RT014Interpretation_Value3r1", "Neg")
830       pubStrRT014Interpretation_Value3r2 = GetOptionSetting("RT014Interpretation_Value3r2", "")
840       pubStrRT014Interpretation_Value4r1 = GetOptionSetting("RT014Interpretation_Value4r1", "Neg")
850       pubStrRT014Interpretation_Value4r2 = GetOptionSetting("RT014Interpretation_Value4r2", "Neg")
860       pubStrRT014Interpretation_Value5r1 = GetOptionSetting("RT014Interpretation_Value5r1", "Neg")
870       pubStrRT014Interpretation_Value5r2 = GetOptionSetting("RT014Interpretation_Value5r2", "Pos")
880       pubStrRT014Interpretation_Value6r1 = GetOptionSetting("RT014Interpretation_Value6r1", "Pos")
890       pubStrRT014Interpretation_Value6r2 = GetOptionSetting("RT014Interpretation_Value6r2", "")
900       pubStrRT014Interpretation_Value7r1 = GetOptionSetting("RT014Interpretation_Value7r1", "Pos")
910       pubStrRT014Interpretation_Value7r2 = GetOptionSetting("RT014Interpretation_Value7r2", "Neg")
920       pubStrRT014Interpretation_Value8r1 = GetOptionSetting("RT014Interpretation_Value8r1", "Pos")
930       pubStrRT014Interpretation_Value8r2 = GetOptionSetting("RT014Interpretation_Value8r2", "Pos")

940   Exit Sub

LoadSpecialTestingSettings_Error:

 Dim strES As String
 Dim intEL As Integer

950    intEL = Erl
960    strES = Err.Description
970    LogError "frmNewStockBar", "LoadSpecialTestingSettings", intEL, strES

End Sub

