Attribute VB_Name = "basFormat"
Option Explicit

Sub OpenPipeDelimitedCsv()

    Dim strPath As String
    
    If InStr(ActiveWorkbook.name, ".csv") = 0 Then
        MsgBox "This is not a csv file", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    ActiveWorkbook.SaveAs Environ("TEMP") & "\" & Replace(ActiveWorkbook.name, ".csv", ".dat")
    
    strPath = ActiveWorkbook.FullName
     
    ActiveWorkbook.Close SaveChanges:=vbNo
    
    Workbooks.Open filename:=strPath, Format:=6, delimiter:=Chr(124)

End Sub

Sub CopySqlInsertStatement()

    Dim sql As String
    Dim col As Integer
    Dim row As Integer
    Dim val As String
    Dim strOriginalStatusBarMsg As String
    
    strOriginalStatusBarMsg = Application.StatusBar
    sql = "INSERT INTO tablename ("
    
    For col = 1 To ActiveSheet.UsedRange.Columns.count
        sql = sql & Cells(1, col) & ","
    Next
    sql = TrimLastCharOff(sql)
    sql = sql & ")" & vbCrLf & "VALUES "
    
    For row = 2 To ActiveSheet.UsedRange.Rows.count
        If Cells(row, 1) = "" Then GoTo Skip
        
        sql = sql & "("
        
        For col = 1 To ActiveSheet.UsedRange.Columns.count
            val = Cells(row, col)
            
            If Not IsNumeric(val) And val <> "NULL" Then val = "'" & val & "'"
            
            sql = sql & val & ","
        Next
        
        sql = TrimLastCharOff(sql)
        
        sql = sql & "),"
        
        If row Mod 1000 = 0 Then Application.StatusBar = "Processing " & row & " records..."
        
        If row <> ActiveSheet.UsedRange.Rows.count Then sql = sql & vbCrLf
Skip:
    Next
    
    sql = TrimLastCharOff(sql)

    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject

    clipboard.SetText sql
    clipboard.PutInClipboard
    
    Application.StatusBar = strOriginalStatusBarMsg
    
End Sub

Function ColLtr(iCol As Long) As String
    If iCol > 0 And iCol <= Columns.count Then
        ColLtr = Evaluate("substitute(address(1, " & iCol & ", 4), ""1"", """")")
    End If
End Function

Sub CopyWithCommas()

    CopySelection ""

End Sub

Sub CopyWithSingleQuotesAndCommas()

    CopySelection "'"

End Sub
Sub CopySelection(ByVal strDelimiter As String)
    
    Dim cell As Variant
    Dim strResult As String
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject

    For Each cell In Selection.SpecialCells(xlVisible)
        If (Len(strResult) = 0) Then
            strResult = strDelimiter & cell & strDelimiter
        Else
            strResult = strResult & "," & strDelimiter & cell & strDelimiter
        End If
    Next

    clipboard.SetText strResult
    clipboard.PutInClipboard
End Sub

Sub FormatDateTime()

    Selection.NumberFormat = "m/d/yyyy h:mm"

End Sub

Sub FormatLzDateTime()
    Dim rng As Range
    
    For Each rng In Selection
        rng.Value = Mid(rng.Value, 8, 2) & "-" & Mid(rng.Value, 6, 2) & "-" & Mid(rng.Value, 1, 4)
    
    Next rng

End Sub

Sub AddQuotesAndCommas()

    AddDecoration """", """"

End Sub

Sub AddSingleQuotesAndCommas()

    AddDecoration "''", "'"

End Sub

Private Sub AddDecoration(ByVal strStartDecorator As String, ByVal strEndDecorator As String)

    Dim rngAppliedRange As Range
    Dim rngLast As Range
    Dim rng As Range
    Dim count As Integer
    Dim result As VbMsgBoxResult
    
    If Selection.count > 1 Then
        Set rngAppliedRange = Selection
    Else
        result = MsgBox("This will apply formatting to the current used range and cannot be undone" & vbCrLf & vbCrLf & "Are you sure you want to continue?", _
                         vbYesNo + vbQuestion, "Used range selection")
        
        If result = vbNo Then Exit Sub
        
        Set rngAppliedRange = ActiveSheet.UsedRange
    End If

    For Each rng In rngAppliedRange
        count = count + 1
        
        If count < rngAppliedRange.count Then
            rng.Value = strStartDecorator & rng.Value & strEndDecorator & ","
        Else
            rng.Value = strStartDecorator & rng.Value & strEndDecorator
        End If
        
    Next rng
    
End Sub

Function GetLastDataCellByRow(ByVal wbkName As String, _
                            ByVal wksName As String) As Excel.Range
'###############################################################################
'   Description : Function to determine last populated cell in the last row of
'                 data in a given worksheet.  This is not the same as
'                 Range("A1").SpecialCells(xlCellTypeLastCell)
'
'   Parameters  : wbkName - Workbook to determine last row from.
'                 wksName - Worksheet to determine last row from.
'
'   Return      : An Excel Range of the last cell in the worksheet containing data.
'###############################################################################

    ' Error Trapping handled by call stack.
    
    Set GetLastDataCellByRow = Workbooks(wbkName).Worksheets(wksName) _
                                    .Cells.Find("*", SearchOrder:=xlByRows, _
                                                SearchDirection:=xlPrevious)
End Function

Function StringFormat(ByRef strText As String, _
                      Optional ByVal strSurrounder As String = "", _
                      Optional ByVal param1 As String = "<empty>", _
                      Optional ByVal param2 As String = "<empty>", _
                      Optional ByVal param3 As String = "<empty>", _
                      Optional ByVal param4 As String = "<empty>", _
                      Optional ByVal param5 As String = "<empty>", _
                      Optional ByVal param6 As String = "<empty>", _
                      Optional ByVal param7 As String = "<empty>", _
                      Optional ByVal param8 As String = "<empty>", _
                      Optional ByVal param9 As String = "<empty>", _
                      Optional ByVal param10 As String = "<empty>", _
                      Optional ByVal param11 As String = "<empty>", _
                      Optional ByVal param12 As String = "<empty>", _
                      Optional ByVal param13 As String = "<empty>", _
                      Optional ByVal param14 As String = "<empty>", _
                      Optional ByVal param15 As String = "<empty>", _
                      Optional ByVal param16 As String = "<empty>", _
                      Optional ByVal param17 As String = "<empty>", _
                      Optional ByVal param18 As String = "<empty>", _
                      Optional ByVal param19 As String = "<empty>") As String

    ' Replaces single quotes with two single quotes (to break it) and line breaks with a space.
    If param1 <> "<empty>" Then strText = Replace(strText, "{0}", strSurrounder & Replace(Replace(param1, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param2 <> "<empty>" Then strText = Replace(strText, "{1}", strSurrounder & Replace(Replace(param2, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param3 <> "<empty>" Then strText = Replace(strText, "{2}", strSurrounder & Replace(Replace(param3, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param4 <> "<empty>" Then strText = Replace(strText, "{3}", strSurrounder & Replace(Replace(param4, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param5 <> "<empty>" Then strText = Replace(strText, "{4}", strSurrounder & Replace(Replace(param5, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param6 <> "<empty>" Then strText = Replace(strText, "{5}", strSurrounder & Replace(Replace(param6, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param7 <> "<empty>" Then strText = Replace(strText, "{6}", strSurrounder & Replace(Replace(param7, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param8 <> "<empty>" Then strText = Replace(strText, "{7}", strSurrounder & Replace(Replace(param8, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param9 <> "<empty>" Then strText = Replace(strText, "{8}", strSurrounder & Replace(Replace(param9, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param10 <> "<empty>" Then strText = Replace(strText, "{9}", strSurrounder & Replace(Replace(param10, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param11 <> "<empty>" Then strText = Replace(strText, "{10}", strSurrounder & Replace(Replace(param11, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param12 <> "<empty>" Then strText = Replace(strText, "{11}", strSurrounder & Replace(Replace(param12, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param13 <> "<empty>" Then strText = Replace(strText, "{12}", strSurrounder & Replace(Replace(param13, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param14 <> "<empty>" Then strText = Replace(strText, "{13}", strSurrounder & Replace(Replace(param14, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param15 <> "<empty>" Then strText = Replace(strText, "{14}", strSurrounder & Replace(Replace(param15, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param16 <> "<empty>" Then strText = Replace(strText, "{15}", strSurrounder & Replace(Replace(param16, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param17 <> "<empty>" Then strText = Replace(strText, "{16}", strSurrounder & Replace(Replace(param17, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param18 <> "<empty>" Then strText = Replace(strText, "{17}", strSurrounder & Replace(Replace(param18, "'", "''"), vbCrLf, " ") & strSurrounder)
    If param19 <> "<empty>" Then strText = Replace(strText, "{18}", strSurrounder & Replace(Replace(param19, "'", "''"), vbCrLf, " ") & strSurrounder)
    
    StringFormat = strText

End Function

Function TrimLastCharOff(ByVal str As String)

    TrimLastCharOff = Left(str, Len(str) - 1)

End Function
