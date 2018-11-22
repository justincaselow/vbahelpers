Attribute VB_Name = "basCompare"
Option Explicit

Public g_blnCompareOptions As CompareOptions

Type CompareOptions
    IncludeColumnHeaders    As Boolean
    IncludeRowHeaders       As Boolean
    HighlightDifferences    As Boolean
    AttachToActiveWorkbook  As Boolean
    CompareAllWorksheets    As Boolean
    OutputWorkbook          As Workbook
    NumericTolerance        As Double
    RowHeaderColumnNum      As Integer
    ColHeaderRowNum         As Integer
    NumToleranceFields      As Variant
End Type

Sub CompareWorkbooks()
    Dim wks1 As Worksheet
    Dim wks2 As Worksheet
    Dim wbkOutput As Workbook
    Dim dctWbk2Sheets As New Dictionary
    
    On Error GoTo ErrorHandler
    
    With frmCompareWorksheets
        .Show
        
        If .Tag <> vbCancel Then
            Application.ScreenUpdating = False
            
            g_blnCompareOptions.IncludeColumnHeaders = .g_blnIncludeColumnHeaders
            g_blnCompareOptions.IncludeRowHeaders = .g_blnIncludeRowHeaders
            g_blnCompareOptions.HighlightDifferences = .g_blnHighlightDifferences
            g_blnCompareOptions.AttachToActiveWorkbook = .g_blnAttachToActiveWorkbook
            g_blnCompareOptions.CompareAllWorksheets = .g_blnCompareAllWorksheets
            g_blnCompareOptions.NumericTolerance = .g_dblNumericTolerance
            g_blnCompareOptions.RowHeaderColumnNum = .g_intRowHeaderColumnNum
            g_blnCompareOptions.ColHeaderRowNum = .g_intColHeaderRowNum
            g_blnCompareOptions.NumToleranceFields = .g_varNumToleranceFields
            
            'Create output worksheet
            Set g_blnCompareOptions.OutputWorkbook = Workbooks.Add
                
            If g_blnCompareOptions.CompareAllWorksheets Then
                'Fill dictionary with worksheets in second workbook
                For Each wks2 In Workbooks(.Workbook2Name).Worksheets
                    dctWbk2Sheets.Add wks2.name, wks2
                Next wks2
                
                For Each wks1 In Workbooks(.Workbook1Name).Worksheets
                    If Not dctWbk2Sheets.Exists(wks1.name) Then
                        MsgBox "Worksheet: " & wks1.name & " from workbook 1 (" & .Workbook1Name & ") does not exist in workbook 2: " & .Workbook2Name
                    Else
                        CompareWorksheets wks1, dctWbk2Sheets(wks1.name), .Workbook1Name, .Workbook2Name
                    End If
                Next wks1
                
                MsgBox "Comparison complete", vbInformation + vbOKOnly, "Complete"
                g_blnCompareOptions.OutputWorkbook.Activate
            Else
                CompareWorksheets .Worksheet1, .Worksheet2, .Workbook1Name, .Workbook2Name
            End If
        End If
        
        Unload frmCompareWorksheets
    End With

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    Application.Cursor = xlDefault
    
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
End Sub
Sub CompareWorksheets(ByVal wks1 As Worksheet, ByVal wks2 As Worksheet, _
                              ByVal Workbook1Name As String, ByVal Workbook2Name As String)

    Dim varUsedRange1 As Variant
    Dim varUsedRange2 As Variant
    Dim row As Long
    Dim col As Long
    Dim colLetter As String
    Dim cellAddress As String
    Dim tablename As String
    Dim strRowHeader As String
    
    Dim wbkResults As Workbook
    Dim wksResults As Worksheet
    Dim dctDifferences As New Dictionary
    Dim rstDifferences As New Recordset
    Dim differenceCount As Long
    Dim lngOffset As Long
    
    Dim cell1 As Variant
    Dim cell2 As Variant
    Dim fieldsArray(6) As Variant
    Dim valuesArray(6) As Variant
    
    If g_blnCompareOptions.HighlightDifferences Then
        wks1.Activate
        wks1.Cells.Select
        Selection.Font.Bold = False
        Selection.Font.Color = vbBlack
        wks1.Range("A1").Select
        wks2.Activate
        wks2.Cells.Select
        Selection.Font.Bold = False
        Selection.Font.Color = vbBlack
        wks2.Range("A1").Select
    End If
    
    Application.Cursor = xlWait
    
    If Len(wks1.Cells(g_blnCompareOptions.ColHeaderRowNum, g_blnCompareOptions.RowHeaderColumnNum).Value) = 0 Then
        strRowHeader = "RowHeader"
    Else
        strRowHeader = wks1.Cells(g_blnCompareOptions.ColHeaderRowNum, g_blnCompareOptions.RowHeaderColumnNum)
    End If
    
    With rstDifferences
        .fields.Append "ColumnHeader", adVariant
        .fields.Append strRowHeader, adVariant
        .fields.Append "Address", adChar, 50 'so that it can be sorted
        .fields.Append "Column", adChar, 50 'so that it can be sorted
        .fields.Append "Row", adVariant
        .fields.Append "Workbook1Value", adVariant
        .fields.Append "Workbook2Value", adVariant
        .Open
    End With
    
    fieldsArray(0) = "ColumnHeader"
    fieldsArray(1) = strRowHeader
    fieldsArray(2) = "Address"
    fieldsArray(3) = "Column"
    fieldsArray(4) = "Row"
    fieldsArray(5) = "Workbook1Value"
    fieldsArray(6) = "Workbook2Value"
    
    varUsedRange1 = wks1.UsedRange
    varUsedRange2 = wks2.UsedRange
    
    Application.StatusBar = "(1 of 2) Checking workbook 1 against workbook 2"
    
    For col = LBound(varUsedRange1, 2) To UBound(varUsedRange1, 2)
        
        colLetter = ColLtr(col)
        
        For row = LBound(varUsedRange1, 1) To UBound(varUsedRange1, 1)
            cellAddress = colLetter & row
            
            cell1 = varUsedRange1(row, col)
            
            If row <= UBound(varUsedRange2) And col <= UBound(varUsedRange2, 2) Then
                cell2 = varUsedRange2(row, col)
            Else
                cell2 = ""
            End If
            
            If IsError(cell1) Then
                cell1 = "#N/A"
            End If
            If IsError(cell2) Then
                cell2 = "#N/A"
            End If
            
            If cell1 <> cell2 Then
                If (IsDate(cell1) And IsDate(cell2)) Then
                    If (CDate(cell1) = CDate(cell2)) Then GoTo ContinueLoop1
                End If
                
                If IsNumeric(cell1) And IsNumeric(cell2) Then
                    If Round(cell1, 12) = Round(cell2, 12) Then GoTo ContinueLoop1
                    If Abs(Round(cell1, 12) - Round(cell2, 12)) < g_blnCompareOptions.NumericTolerance And _
                       IsInArray(colLetter, g_blnCompareOptions.NumToleranceFields) Then GoTo ContinueLoop1
                End If
                
                differenceCount = differenceCount + 1
                
                dctDifferences.Add cellAddress, _
                    cellAddress & " on " & Workbook1Name & " has value (" & cell1 & ") on " & _
                    Workbook2Name & " and has value (" & cell2 & ")"
                    
                valuesArray(0) = Cells(g_blnCompareOptions.ColHeaderRowNum, col).Value
                valuesArray(1) = Cells(row, g_blnCompareOptions.RowHeaderColumnNum).Value
                valuesArray(2) = cellAddress
                valuesArray(3) = colLetter
                valuesArray(4) = row
                valuesArray(5) = cell1
                valuesArray(6) = cell2
                
                rstDifferences.AddNew fieldsArray, valuesArray
            End If
ContinueLoop1:
        Next
    Next
        
    Application.StatusBar = "(2 of 2) Outputting results"
    
    If differenceCount = 0 Then
        If Not g_blnCompareOptions.CompareAllWorksheets Then
            'Only show dialog if compare all worksheets is not checked
            MsgBox "No differences found", vbInformation + vbOKOnly, "No differences"
        End If
    Else
        If Not g_blnCompareOptions.OutputWorkbook Is Nothing Then
            Set wbkResults = g_blnCompareOptions.OutputWorkbook
            Set wksResults = wbkResults.ActiveSheet
        ElseIf g_blnCompareOptions.AttachToActiveWorkbook Then
            Set wbkResults = ActiveWorkbook
            Set wksResults = wbkResults.Worksheets.Add(After:=wbkResults.ActiveSheet)
        Else
            Set wbkResults = Workbooks.Add
            Set wksResults = wbkResults.ActiveSheet
        End If
        
        With wksResults
            .Activate
            
            If g_blnCompareOptions.OutputWorkbook Is Nothing Then
                .name = "Differences_" & Left(wks1.name, 19)
            End If
            
            If GetLastDataCellByRow(wbkResults.name, wksResults.name) Is Nothing Then
                lngOffset = 0
            Else
                lngOffset = GetLastDataCellByRow(wbkResults.name, wksResults.name).row + 2
            End If

            .Range("A1").Offset(lngOffset, 0) = "Workbook 1 is " & Workbook1Name & " (Worksheet Name: " & wks1.name & ")"
            .Range("A2").Offset(lngOffset, 0) = "Workbook 2 is " & Workbook2Name & " (Worksheet Name: " & wks2.name & ")"
            .Range("A3").Offset(lngOffset, 0) = "Comparison run:"
            .Range("B3").Offset(lngOffset, 0) = Format(Now, "dd-mmm-yyyy HH:mm:ss")
            
            rstDifferences.MoveFirst
            'By default sort by address
            rstDifferences.Sort = rstDifferences.fields(2).name
            
            .Range("A4").Offset(lngOffset, 0).Value = rstDifferences.fields(0).name
            .Range("B4").Offset(lngOffset, 0).Value = rstDifferences.fields(1).name
            .Range("C4").Offset(lngOffset, 0).Value = rstDifferences.fields(2).name
            .Range("D4").Offset(lngOffset, 0).Value = rstDifferences.fields(3).name
            .Range("E4").Offset(lngOffset, 0).Value = rstDifferences.fields(4).name
            .Range("F4").Offset(lngOffset, 0).Value = rstDifferences.fields(5).name & " (" & wks1.name & ")"
            .Range("G4").Offset(lngOffset, 0).Value = rstDifferences.fields(6).name & " (" & wks2.name & ")"
            .Range("H4").Offset(lngOffset, 0).Value = "Difference"
             
            '.Cells.Rows(4).Font.Bold = True
            .Range("A5").Offset(lngOffset, 0).CopyFromRecordset rstDifferences
            
            If g_blnCompareOptions.HighlightDifferences Then
                rstDifferences.MoveFirst
                
                While Not rstDifferences.EOF
                    wks1.Range(rstDifferences.fields(2).Value).Font.Bold = True
                    wks1.Range(rstDifferences.fields(2).Value).Font.Color = vbRed
                    wks2.Range(rstDifferences.fields(2).Value).Font.Bold = True
                    wks2.Range(rstDifferences.fields(2).Value).Font.Color = vbRed
                    rstDifferences.MoveNext
                Wend
            End If
            
            'Format table
            tablename = "Table_" & Format(Now, "yyyymmdd_HHmmss_") & Math.Rnd() * 10000
            .Range("A4").Select
            .ListObjects.Add(xlSrcRange, .Range( _
                        .Range("A4").Offset(lngOffset, 0), _
                        .Range("H" & (4 + rstDifferences.RecordCount)).Offset(lngOffset, 0)), , xlYes).name = tablename
            .Range(tablename).Select
            .ListObjects(tablename).TableStyle = "TableStyleMedium16"
            .Range("H5").Offset(lngOffset, 0).Formula = "=IFERROR(" & Range("G5").Offset(lngOffset, 0).Address(False, False) & "-" & Range("F5").Offset(lngOffset, 0).Address(False, False) & ", """")"
            
            .Range("F:H").NumberFormat = "#,##0_ ;[Red]-#,##0 "
            
            .Cells.Select
            .Cells.EntireColumn.AutoFit
            .Columns(1).ColumnWidth = 50
            
            .Range("A1:A2").Offset(lngOffset, 0).Style = "Heading 4"
            .Range("A1").Select
            
            If Not g_blnCompareOptions.CompareAllWorksheets Then
                MsgBox differenceCount & " differences found", vbInformation + vbOKOnly, "Differences found"
            End If
        End With
        
    End If
    
    Application.StatusBar = ""

ExitSub:

    Application.StatusBar = ""
    Application.Cursor = xlDefault
    
End Sub
