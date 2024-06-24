VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGraph 
   BackColor       =   &H00E0E0E0&
   Caption         =   "NetAcquire --- Graph"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14880
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShowLegend 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show Legend"
      Height          =   255
      Left            =   3420
      TabIndex        =   11
      Top             =   240
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   7635
      Left            =   120
      OleObjectBlob   =   "frmGraph.frx":0000
      TabIndex        =   1
      Top             =   180
      Width           =   14595
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   8415
      Left            =   16800
      TabIndex        =   5
      Top             =   600
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   14843
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   60
      ScaleHeight     =   1425
      ScaleWidth      =   14670
      TabIndex        =   0
      Top             =   7920
      Width           =   14700
      Begin VB.OptionButton optGraphType 
         Caption         =   "3D Area"
         Height          =   1100
         Index           =   5
         Left            =   6420
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   1200
      End
      Begin VB.OptionButton optGraphType 
         Caption         =   "3D Line"
         Height          =   1100
         Index           =   4
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1200
      End
      Begin VB.OptionButton optGraphType 
         Caption         =   "3D Bar"
         Height          =   1100
         Index           =   3
         Left            =   3900
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   1200
      End
      Begin VB.OptionButton optGraphType 
         Caption         =   "2D Area"
         Height          =   1100
         Index           =   2
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1200
      End
      Begin VB.OptionButton optGraphType 
         Caption         =   "2D Line"
         Height          =   1100
         Index           =   1
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1200
      End
      Begin VB.OptionButton optGraphType 
         Caption         =   "2D Bar"
         Height          =   1100
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   1100
         Left            =   13320
         TabIndex        =   4
         Top             =   180
         Width           =   1200
      End
      Begin VB.CommandButton cmdRotate 
         Caption         =   "Rotate"
         Enabled         =   0   'False
         Height          =   1100
         Left            =   10680
         TabIndex        =   3
         Top             =   180
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   1100
         Left            =   11940
         TabIndex        =   2
         Top             =   180
         Width           =   1200
      End
   End
   Begin MSComDlg.CommonDialog dlgChart 
      Left            =   120
      Top             =   7980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Programmer Mohammad Hassan Nolehdan
' Nolehdan@yahoo.com
' Please contact with me for more information or upgrad version
' I will glad if you correct me for buge

Dim MyTable As String
Dim MyField1 As String
Dim MyField2 As String
Dim MyField3 As String
Dim MyText1 As String
Dim MyText2 As String
Dim MyText3 As String
Dim ii As Double
Dim e As Double
Dim R As Double
Dim n As Integer
Dim n1 As Integer
Dim n2 As Integer
Dim n3 As Integer
Dim n4 As Integer
Dim n5 As Integer
Dim n6 As Integer
Dim i As Integer
Dim J As Integer
Dim ColNum As Integer
Dim Recno As Integer
Dim Rl(100) As String
Dim Cl(100) As String

Private m_sGraphTitleText As String
Private m_sGraphFootNoteText As String

Private Sub DrawChart()
          Dim iRow As Integer
          Dim iCol As Integer
          Dim FontName As String
          Dim MinValue As Double
          Dim MaxValue As Double
          
63530     On Error GoTo DrawChart_Error

63540     FontName = "Courier New"
63550     With MSChart1
              '*******LOOK AND FEEL
63560         .RowCount = g.Rows - 1
63570         .ColumnCount = g.Cols - 1
63580         .Backdrop.Fill.Style = VtFillStyleBrush
63590         .Backdrop.Fill.Brush.FillColor.Set 0, 200, 250
63600         .Title.VtFont.size = 10
63610         .Title.VtFont.Name = FontName
63620         .Footnote.VtFont.size = 9
63630         .Footnote.VtFont.Name = FontName
63640         .TitleText = m_sGraphTitleText
63650         .FootnoteText = m_sGraphFootNoteText
              
63660         For iCol = 1 To .ColumnCount
63670             For iRow = 1 To .RowCount
63680                 Debug.Print g.TextMatrix(iRow, iCol)
63690                 If CDbl(g.TextMatrix(iRow, iCol)) < MinValue Then
63700                     MinValue = CDbl(g.TextMatrix(iRow, iCol))
63710                 End If
63720                 If CDbl(g.TextMatrix(iRow, iCol)) > MaxValue Then
63730                     MaxValue = CDbl(g.TextMatrix(iRow, iCol))
63740                 End If
63750             Next iRow
63760         Next iCol
63770         If MaxValue Mod 10 > 0 Then
63780             MaxValue = MaxValue + (10 - (MaxValue Mod 10))
63790         End If
              
63800         .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Name = FontName
63810         .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.size = 9
63820         .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
63830         .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinValue
63840         .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxValue
63850         .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 10
              '.Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0
              
63860         For iCol = 1 To .ColumnCount
63870             With .Plot.SeriesCollection(iCol)
63880                 .LegendText = g.TextMatrix(0, iCol)
                     
63890             End With
                  
63900             For iRow = 1 To .RowCount
63910                 .row = iRow
63920                 .RowLabel = g.TextMatrix(iRow, 0)
63930                 .DataGrid.SetData iRow, iCol, CDbl(g.TextMatrix(iRow, iCol)), False
63940                 If CDbl(g.TextMatrix(iRow, iCol)) > MinValue Then
63950                     MinValue = CDbl(g.TextMatrix(iRow, iCol))
63960                 End If
63970                 If MaxValue < CDbl(g.TextMatrix(iRow, iCol)) Then
63980                     MaxValue = CDbl(g.TextMatrix(iRow, iCol))
63990                 End If
                      
64000             Next iRow
64010         Next iCol
              
              
              
              

64020     End With



64030     Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

64040     intEL = Erl
64050     strES = Err.Description
64060     LogError "frmGraph", "DrawChart", intEL, strES


End Sub
Private Sub DrawChartFromArray()
          Dim iRow As Integer
          Dim iCol As Integer
          Dim arrSeries()
64070     On Error GoTo DrawChartFromArray_Error

64080     With MSChart1
              '*******LOOK AND FEEL

64090         ReDim arrSeries(1 To g.Rows - 1, 1 To g.Cols - 1)
64100         .Backdrop.Fill.Style = VtFillStyleBrush
64110         .Backdrop.Fill.Brush.FillColor.Set 0, 200, 250
64120         .Title.VtFont.size = 10
64130         .Title.VtFont.Name = "Courier New"

64140         .Footnote.VtFont.size = 9
64150         .Footnote.VtFont.Name = "Courier New"

64160         .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Name = "Courier New"
64170         .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.size = 9

64180         .RowCount = g.Rows - 1
64190         .ColumnCount = g.Cols - 1
64200         .TitleText = m_sGraphTitleText
64210         .FootnoteText = m_sGraphFootNoteText

64220         For iCol = 1 To .ColumnCount
64230             For iRow = 1 To .RowCount
64240                 arrSeries(iRow, iCol) = g.TextMatrix(iRow, iCol)
64250             Next iRow
64260         Next iCol
64270         .ChartData = arrSeries

              '    For iCol = 1 To .ColumnCount
              '        .Plot.SeriesCollection(iCol).LegendText = g.TextMatrix(0, iCol)
              '
              '        For iRow = 1 To .RowCount
              '            .row = iRow
              '            .RowLabel = g.TextMatrix(iRow, 0)
              '            .DataGrid.SetData iRow, iCol, CDbl(g.TextMatrix(iRow, iCol)), False
              '        Next iRow
              '    Next iCol



64280     End With



64290     Exit Sub

DrawChartFromArray_Error:

          Dim strES As String
          Dim intEL As Integer

64300     intEL = Erl
64310     strES = Err.Description
64320     LogError "frmGraph", "DrawChartFromArray", intEL, strES


End Sub


'Private Sub chkInvers_Click()
'If chkInvers.Value = 1 Then
'    MSChart1.Plot.DataSeriesInRow = True
'Else
'    MSChart1.Plot.DataSeriesInRow = False
'End If
'End Sub

Private Sub chkShowLegend_Click()
64330     On Error GoTo chkShowLegend_Click_Error

64340     If chkShowLegend.Value = 1 Then
64350         MSChart1.ShowLegend = True
64360     Else
64370         MSChart1.ShowLegend = False
64380     End If

64390     Exit Sub

chkShowLegend_Click_Error:

          Dim strES As String
          Dim intEL As Integer

64400     intEL = Erl
64410     strES = Err.Description
64420     LogError "frmGraph", "chkShowLegend_Click", intEL, strES

End Sub

'Private Sub Combo2_Click()
'
'On Error Resume Next
'
'MSChart1.SeriesType = Combo2.ListIndex
'
'End Sub




Private Sub cmdPrint_Click()

64430     On Error GoTo cmdPrint_Click_Error

64440     On Error Resume Next

64450     MSChart1.EditCopy
64460     Printer.Orientation = vbPRORLandscape
64470     Printer.Print " "
64480     Printer.PaintPicture Clipboard.GetData(), 0, 0, 15000, 11000
64490     Printer.EndDoc


64500     Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

64510     intEL = Erl
64520     strES = Err.Description
64530     LogError "frmGraph", "cmdPrint_Click", intEL, strES


End Sub

Private Sub cmdRotate_Click()

64540     On Error GoTo cmdRotate_Click_Error

64550     On Error Resume Next

64560     e = MSChart1.Plot.View3d.Elevation
64570     R = MSChart1.Plot.View3d.Rotation

64580     MSChart1.Plot.View3d.Set R + ii, e
64590     ii = ii + 1

64600     Exit Sub

cmdRotate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

64610     intEL = Erl
64620     strES = Err.Description
64630     LogError "frmGraph", "cmdRotate_Click", intEL, strES


End Sub

'Private Sub emdOpen_Click()
'
'On Error Resume Next
'
'Dim a As Variant
'Dim b As Variant
'Dim tt As String
'
'CommonDialog1.Filter = "Graph Format(*.Grc)|*.Grc|"
'
'CommonDialog1.ShowOpen
'
'Call ClearData
'
'Open CommonDialog1.FileName For Input As #1
'
'Input #1, a, b
'If a = "ChartType" Then
'    MSChart1.chartType = b
'End If
'
'Input #1, a, b
'If a = "UpName" Then
'    Text2.Text = b
'End If
'
'Input #1, a, b
'If a = "DownName" Then
'    Text3.Text = b
'End If
'
'Input #1, a, b
'If a = "AvamelNo" Then
'    Text1.Text = b
'    MSChart1.ColumnCount = b
'End If
'
'Input #1, a, b
'If a = "RecordNo" Then
'    Recno = b
'    MSChart1.RowCount = b
'End If
'
'For i = 1 To MSChart1.RowCount
'    Input #1, a, b
'    If a = "RowName" Then
'        'MSChart1.row = i
'        'MSChart1.RowLabel = b
'        Rl(i) = b
'    End If
'Next i
'
'For i = 1 To MSChart1.ColumnCount
'    Input #1, a, b
'    If a = "ColName" Then
'        'MSChart1.column = i
'        'MSChart1.ColumnLabel = b
'        Cl(i) = b
'    End If
'Next i
'
'Adodc1.Refresh
'For i = 1 To Recno
'    Adodc1.Recordset.AddNew
'    For j = 1 To Val(Text1.Text)
'        Input #1, a, b
'        Adodc1.Recordset(Adodc1.Recordset.Fields(j - 1).Name).Value = b
'    Next j
'    Adodc1.Recordset.Update
'    Adodc1.Recordset.MoveNext
'Next i
'
'Close #1
'
'For i = 1 To MSChart1.RowCount
'    MSChart1.row = i
'    MSChart1.RowLabel = Rl(i)
'Next i
'
'For i = 1 To MSChart1.ColumnCount - 1
'    MSChart1.column = i
'    MSChart1.ColumnLabel = Cl(i)
'    DataGrid1.Columns(i - 1).Caption = Cl(i)
'Next i
'
'End Sub

'Private Sub cmdSave_Click()
'
'On Error Resume Next
'
'CommonDialog1.Filter = "Graph Format(*.Grc)|*.Grc|"
'
'CommonDialog1.ShowSave
'
'If CommonDialog1.CancelError = True Then Exit Sub
'
'Open CommonDialog1.FileName For Output As #1
'
'Write #1, "ChartType", MSChart1.chartType
'Write #1, "UpName", Text2.Text
'Write #1, "DownName", Text3.Text
'Write #1, "AvamelNo", Adodc1.Recordset.Fields.Count
'Write #1, "RecordNo", Adodc1.Recordset.RecordCount
'
'For i = 1 To MSChart1.RowCount
'    MSChart1.row = i
'    Write #1, "RowName", MSChart1.RowLabel
'Next i
'For i = 1 To MSChart1.ColumnCount
'    MSChart1.column = i
'    Write #1, "ColName", MSChart1.ColumnLabel
'Next i
'
'Adodc1.Refresh
'For i = 1 To Adodc1.Recordset.RecordCount
'    For j = 1 To Adodc1.Recordset.Fields.Count
'        Write #1, "RecordTxt", Adodc1.Recordset.Fields(j - 1).Value
'    Next j
'    Adodc1.Recordset.MoveNext
'Next i
'
'Close #1
'
'For i = 1 To MSChart1.ColumnCount
'    MSChart1.column = i
'    DataGrid1.Columns(i - 1).Caption = MSChart1.ColumnLabel
'Next i
'
'End Sub


Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)

          'MsgBox (DataError & "   " & Response) '6153

64640     On Error GoTo DataGrid1_Error_Error

64650     Response = 0

64660     Exit Sub

DataGrid1_Error_Error:

          Dim strES As String
          Dim intEL As Integer

64670     intEL = Erl
64680     strES = Err.Description
64690     LogError "frmGraph", "DataGrid1_Error", intEL, strES


End Sub


Private Sub FillSample()

          Dim Column As Integer
          Dim row As Integer
          Dim index1 As Integer
          Dim index2 As Integer
          Dim index3 As Integer
          Dim index4 As Integer


64700     On Error GoTo FillSample_Error

64710     With MSChart1
              ' Displays a 3d chart with 8 columns and 8 rows
              ' data.
64720         .chartType = VtChChartType3dBar
64730         .ColumnCount = 8
64740         .RowCount = 8
64750         For Column = 1 To 8
64760             For row = 1 To 8
64770                 .Column = Column
64780                 .row = row
64790                 .Data = row * 10
64800             Next row
64810         Next Column
              ' Use the chart as the backdrop of the legend.
64820         .ShowLegend = True
64830         .SelectPart VtChPartTypePlot, index1, index2, _
                  index3, index4
64840         .EditCopy
64850         .SelectPart VtChPartTypeLegend, index1, _
                  index2, index3, index4
64860         .EditPaste
64870     End With

64880     Exit Sub

FillSample_Error:

          Dim strES As String
          Dim intEL As Integer

64890     intEL = Erl
64900     strES = Err.Description
64910     LogError "frmGraph", "FillSample", intEL, strES


End Sub




Private Sub cmdExit_Click()
64920     Unload Me
End Sub

Private Sub Form_Activate()


64930     On Error GoTo Form_Activate_Error
64940     DrawChart


64950     Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

64960     intEL = Erl
64970     strES = Err.Description
64980     LogError "frmGraph", "Form_Activate", intEL, strES

64990     Exit Sub



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

65000     On Error GoTo Form_KeyDown_Error

65010     If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Call cmdRotate_Click

65020     Exit Sub

Form_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

65030     intEL = Erl
65040     strES = Err.Description
65050     LogError "frmGraph", "Form_KeyDown", intEL, strES


End Sub

Private Sub Form_Load()

65060     On Error GoTo Form_Load_Error

65070     On Error Resume Next



          'Call FillMschartColor


          'MSChart1.Backdrop.Shadow.Style = VtShadowStyleDrop
          'MSChart1.Backdrop.Frame.Style = VtFrameStyleSingleLine
          'MSChart1.Backdrop.Fill.Style = VtFillStyleBrush
          'MSChart1.Plot.DepthToHeightRatio = 300
          'MSChart1.DrawMode = VtChDrawModeDraw

65080     MSChart1.DoSetCursor = True
65090     MSChart1.ShowLegend = True
65100     MSChart1.chartType = VtChChartType2dCombination




65110     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

65120     intEL = Erl
65130     strES = Err.Description
65140     LogError "frmGraph", "Form_Load", intEL, strES


End Sub




Private Sub FillMschartColor()

65150     On Error GoTo FillMschartColor_Error

65160     On Error Resume Next

65170     With MSChart1.Backdrop.Fill
65180         .Style = VtFillStyleBrush
65190         .Brush.FillColor.Set 0, 200, 250    ' Set the color of the fill.
65200     End With


65210     Exit Sub

FillMschartColor_Error:

          Dim strES As String
          Dim intEL As Integer

65220     intEL = Erl
65230     strES = Err.Description
65240     LogError "frmGraph", "FillMschartColor", intEL, strES


End Sub

Private Sub MSChart1_KeyDown(KeyCode As Integer, Shift As Integer)

65250     On Error Resume Next

65260     If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then Call cmdRotate_Click

End Sub



Private Sub MSChart1_SeriesActivated(Series As Integer, MouseFlags As Integer, Cancel As Integer)

          ' The CommonDialog control is named dlgChart.
          Dim red, green, blue As Integer
65270     On Error GoTo MSChart1_SeriesActivated_Error

65280     With dlgChart    ' CommonDialog object
65290         .CancelError = True
65300         .ShowColor
65310         red = RedFromRGB(.Color)
65320         green = GreenFromRGB(.Color)
65330         blue = BlueFromRGB(.Color)
65340     End With

          ' NOTE: Only the 2D and 3D line charts use the
          ' Pen object. All other types use the Brush.

65350     If MSChart1.chartType <> VtChChartType2dLine Or _
              MSChart1.chartType <> VtChChartType3dLine Then
65360         MSChart1.Plot.SeriesCollection(Series). _
                  DataPoints(-1).Brush.FillColor. _
                  Set red, green, blue
65370     Else
65380         MSChart1.Plot.SeriesCollection(Series).Pen.VtColor.Set red, green, blue
65390     End If
65400     Exit Sub

65410     Exit Sub

MSChart1_SeriesActivated_Error:

          Dim strES As String
          Dim intEL As Integer

65420     intEL = Erl
65430     strES = Err.Description
65440     LogError "frmGraph", "MSChart1_SeriesActivated", intEL, strES


End Sub
' Paste these functions into the Declarations section
' of the Form or Code Module.
Public Function RedFromRGB(ByVal rgb As Long) As Integer

65450     On Error Resume Next

          ' The ampersand after &HFF coerces the number as a
          ' long, preventing Visual Basic from evaluating the
          ' number as a negative value. The logical And is
          ' used to return bit values.
65460     RedFromRGB = &HFF& And rgb

End Function

Public Function GreenFromRGB(ByVal rgb As Long) As Integer

65470     On Error Resume Next

          ' The result of the And operation is divided by
          ' 256, to return the value of the middle bytes.
          ' Note the use of the Integer divisor.
65480     GreenFromRGB = (&HFF00& And rgb) \ 256

End Function

Public Function BlueFromRGB(ByVal rgb As Long) As Integer

65490     On Error Resume Next

          ' This function works like the GreenFromRGB above,
          ' except you don't need the ampersand. The
          ' number is already a long. The result divided by
          ' 65536 to obtain the highest bytes.
65500     BlueFromRGB = (&HFF0000 And rgb) \ 65536

End Function









Public Property Get GraphTitleText() As String

65510     GraphTitleText = m_sGraphTitleText

End Property

Public Property Let GraphTitleText(ByVal sGraphTitleText As String)

65520     m_sGraphTitleText = sGraphTitleText

End Property

Public Property Get GraphFootNoteText() As String

65530     GraphFootNoteText = m_sGraphFootNoteText

End Property

Public Property Let GraphFootNoteText(ByVal sGraphFootNoteText As String)

10        m_sGraphFootNoteText = sGraphFootNoteText

End Property


Private Sub optGraphType_Click(Index As Integer)

20        On Error GoTo optGraphType_Click_Error

30        Select Case Index
              Case 0: MSChart1.chartType = VtChChartType2dBar
40                cmdRotate.Enabled = False
50            Case 1: MSChart1.chartType = VtChChartType2dLine
60                cmdRotate.Enabled = False
70            Case 2: MSChart1.chartType = VtChChartType2dArea
80                cmdRotate.Enabled = False
90            Case 3: MSChart1.chartType = VtChChartType3dBar
100               cmdRotate.Enabled = True
110           Case 4: MSChart1.chartType = VtChChartType3dLine
120               cmdRotate.Enabled = True
130           Case 5: MSChart1.chartType = VtChChartType3dArea
140               cmdRotate.Enabled = True
150       End Select

160       Exit Sub

optGraphType_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmGraph", "optGraphType_Click", intEL, strES

End Sub
