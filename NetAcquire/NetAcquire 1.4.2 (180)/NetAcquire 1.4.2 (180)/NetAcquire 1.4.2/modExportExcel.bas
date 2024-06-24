Attribute VB_Name = "modExportExcel"
Option Explicit

Public Sub ExportFlexGrid(ByVal objGrid As MSFlexGrid, _
                          ByVal CallingForm As Form, _
                          Optional ByVal HeadingMatrix As String = "")


      Dim objXL As Object
      Dim objWB As Object
      Dim objWS As Object
      Dim R As Long
      Dim c As Long
      Dim t As Single

      'Assume the calling form has a MSFlexGrid (grdToExport),
      'CommandButton (cmdXL) and Label (lblExcelInfo) (Visible set to False)
      'In the calling form:
      'Private Sub cmdXL_Click()
      'ExportFlexGrid grdToExport, Me
      'End Sub

690   On Error GoTo ehEFG

700   With CallingForm.lblExcelInfo
710     .Caption = "Exporting..."
720     .Visible = True
730     .Refresh
740   End With

750   Set objXL = CreateObject("Excel.Application")
760   Set objWB = objXL.Workbooks.Add
770   Set objWS = objWB.Worksheets(1)

      Dim intLineCount As Integer
      '****Change: Babar Shahzad 2007-11-19
      'Heading for export to excel can be passed as string which would be
      'a string having TABS as column breaks and CR as row break.

780   intLineCount = 0
790   If HeadingMatrix <> "" Then
800       With objWS
              Dim strTokens() As String
810           strTokens = Split(HeadingMatrix, vbCr)
820           intLineCount = UBound(strTokens)
        
830           For R = LBound(strTokens) To UBound(strTokens) - 1
                  'For C = 0 To objGrid.Cols - 1
                'The "'" is required to format the cells as text in Excel
                'otherwise entries like "4/2" are interpreted as a date
840             .Range(.Cells(R + 1, 1), .Cells(R + 1, objGrid.Cols)).MergeCells = True
850             .Range(.Cells(R + 1, 1), .Cells(R + 1, objGrid.Cols)).HorizontalAlignment = 3
860             .Range(.Cells(R + 1, 1), .Cells(R + 1, objGrid.Cols)).Font.Bold = True
870             objWS.Cells(R + 1, 1) = "'" & strTokens(R)
          
880           Next
890       End With
         
900   End If

910   With objWS
920       For R = 0 To objGrid.Rows - 1
930           For c = 0 To objGrid.Cols - 1
                  'The "'" is required to format the cells as text in Excel
                  'otherwise entries like "4/2" are interpreted as a date
940               If R = 0 Then
950                   .Range(.Cells(R + 1 + intLineCount, 1), .Cells(R + 1 + intLineCount, objGrid.Cols)).Font.Bold = True
          
960               End If
970               .Cells(R + 1 + intLineCount, c + 1) = "'" & objGrid.TextMatrix(R, c)
980           Next
990       Next
          
1000      .Cells.Columns.AutoFit
1010  End With

1020  objXL.Visible = True

1030  Set objWS = Nothing
1040  Set objWB = Nothing
1050  Set objXL = Nothing

1060  CallingForm.lblExcelInfo.Visible = False

1070  Exit Sub

ehEFG:
      Dim er As Long
      Dim es As String

1080  er = Err.Number
1090  es = Err.Description

1100  iMsg es

1110  With CallingForm.lblExcelInfo
1120    .Caption = "Error " & Format(er)
1130    .Refresh
1140    t = Timer
1150    Do While Timer - t < 1: Loop
1160    .Visible = False
1170  End With

1180  Exit Sub

      '
      '
      '
      '
      '
      '
      '
      '
      '
      '
      '
      '
      '      Dim objXL As Object
      '      Dim objWB As Object
      '      Dim objWS As Object
      '      Dim r As Long
      '      Dim C As Long
      '      Dim t As Single
      '      'Assume the calling form has a MSFlexGrid (grdToExport),
      '      'CommandButton (cmdXL) and Label (lblExcelInfo) (Visible set to False)
      '      'In the calling form:
      '      'Private Sub cmdXL_Click()
      '      'ExportFlexGrid grdToExport, Me
      '      'End Sub
      '
      '
      '10    On Error GoTo ehEFG
      '
      '20    With CallingForm.lblExcelInfo
      '30      .Caption = "Exporting..."
      '40      .Visible = True
      '50      .Refresh
      '60    End With
      '
      '70    Set objXL = CreateObject("Excel.Application")
      '80    Set objWB = objXL.Workbooks.Add
      '90    Set objWS = objWB.Worksheets(1)
      '
      '100   With objWS
      '110     For r = 0 To objGrid.Rows - 1
      '120       For C = 0 To objGrid.Cols - 1
      '130         If C < 256 Then '0-255 max columns in Excel
      '              'The "'" is required to format the cells as text in Excel
      '              'otherwise entries like "4/2" are interpreted as a date
      '140           .Cells(r + 1, C + 1) = "'" & objGrid.TextMatrix(r, C)
      '150         End If
      '160       Next
      '170     Next
      '
      '180     .Cells.Columns.AutoFit
      '190   End With
      '
      '      'This line removes the green error triangle
      '      'Does not work with XL2000!!! - Requires at least XL2002
      '      'objXL.ErrorCheckingOptions.NumberAsText = False
      '
      '200   objXL.Visible = True
      '
      '210   Set objWS = Nothing
      '220   Set objWB = Nothing
      '230   Set objXL = Nothing
      '
      '240   CallingForm.lblExcelInfo.Visible = False
      '
      '250   Exit Sub
      '
      'ehEFG:
      '      Dim er As Long
      '      Dim es As String
      '
      '260   er = Err.Number
      '270   es = Err.Description
      '
      '280   iMsg es
      '
      '290   With CallingForm.lblExcelInfo
      '300     .Caption = "Error " & Format(er)
      '310     .Refresh
      '320     t = Timer
      '330     Do While Timer - t < 1: Loop
      '340     .Visible = False
      '350   End With
      '
      '360   Exit Sub

End Sub


