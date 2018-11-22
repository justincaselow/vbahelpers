VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompareWorksheets 
   Caption         =   "Compare worksheets"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   OleObjectBlob   =   "frmCompareWorksheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCompareWorksheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Worksheet1 As Worksheet
Public Worksheet2 As Worksheet
Public Workbook1Name As String
Public Workbook2Name As String
Public g_blnIncludeRowHeaders As Boolean
Public g_blnIncludeColumnHeaders As Boolean
Public g_blnHighlightDifferences As Boolean
Public g_blnAttachToActiveWorkbook As Boolean
Public g_blnCompareAllWorksheets As Boolean
Public g_dblNumericTolerance As Double
Public g_intRowHeaderColumnNum As Integer
Public g_intColHeaderRowNum As Integer
Public g_varNumToleranceFields As Variant

Private Sub chkCompareAllWorksheets_Click()

    If chkCompareAllWorksheets.Value Then
        lstWorksheet1Names.Visible = False
        lstWorksheet2Names.Visible = False
    Else
        lstWorksheet1Names.Visible = True
        lstWorksheet2Names.Visible = True
    End If

End Sub

Private Sub chkNumericTolerance_Click()
    txtNumericTolerance.Enabled = chkNumericTolerance.Value
End Sub

Private Sub CommandButton1_Click()
    If Worksheet1 Is Nothing Or Worksheet2 Is Nothing Or Len(Workbook1Name) = 0 Or Len(Workbook2Name) = 0 Then
        MsgBox "Ensure an item is selected from every listbox!", vbExclamation + vbOKOnly, "Validation error"
    Else
        g_blnIncludeColumnHeaders = chkIncludeColumnHeaders.Value
        g_blnIncludeRowHeaders = chkIncludeRowHeaders.Value
        g_blnHighlightDifferences = chkHighlightDifferences.Value
        g_blnAttachToActiveWorkbook = chkAttachToActiveWorkbook.Value
        g_blnCompareAllWorksheets = chkCompareAllWorksheets.Value
        g_intRowHeaderColumnNum = txtRowHeaderColumnNum.Value
        g_intColHeaderRowNum = txtColHeaderRowNum.Value
        g_varNumToleranceFields = Split(txtColumnsForNumericTolerance.Text, ",")
        
        If Len(txtNumericTolerance.Value) > 0 Then
            g_dblNumericTolerance = CDbl(txtNumericTolerance.Value)
        End If
        
        Me.Tag = vbOK
        Me.Hide
    End If
End Sub

Private Sub CommandButton2_Click()
    
    Me.Tag = vbCancel
    Me.Hide
        
End Sub

Private Sub lstWorkbook1Names_Click()

    Workbook1Name = lstWorkbook1Names.List(lstWorkbook1Names.ListIndex)
    ListWorksheets lstWorksheet1Names, Workbook1Name
        
End Sub

Private Sub lstWorkbook2Names_Click()

    Workbook2Name = lstWorkbook2Names.List(lstWorkbook2Names.ListIndex)
    ListWorksheets lstWorksheet2Names, Workbook2Name

End Sub

Private Sub lstWorksheet1Names_Click()
    Set Worksheet1 = Workbooks(Workbook1Name).Worksheets(lstWorksheet1Names.List(lstWorksheet1Names.ListIndex))
End Sub

Private Sub lstWorksheet2Names_Click()
    Set Worksheet2 = Workbooks(Workbook2Name).Worksheets(lstWorksheet2Names.List(lstWorksheet2Names.ListIndex))
End Sub

Private Sub UserForm_Initialize()

    Dim wbk As Workbook
    Dim fso As New FileSystemObject
    Dim strCodeLocation As String
    Dim intIndex As Integer
    
    On Error GoTo ErrorHandler
    
    For Each wbk In Workbooks
        strCodeLocation = "Getting file name for " & wbk.name
        
        lstWorkbook1Names.AddItem wbk.name
        lstWorkbook2Names.AddItem wbk.name
        
        If ActiveWorkbook.name = wbk.name Then
            lstWorkbook1Names.Selected(intIndex) = True
            lstWorkbook2Names.Selected(intIndex) = True
            
            lstWorksheet1Names.Selected(0) = True
            If lstWorksheet2Names.ListCount > 1 Then
                lstWorksheet2Names.Selected(1) = True
            End If
        End If
        
        intIndex = intIndex + 1
    Next
    
    Exit Sub
ErrorHandler:
    MsgBox strCodeLocation & vbCrLf & vbCrLf & Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub ListWorksheets(ByRef lstListBox As MSForms.ListBox, ByVal strWorkbookFileName As String)

    lstListBox.Clear
    Dim wks As Worksheet
    
    For Each wks In Workbooks(strWorkbookFileName).Worksheets
        lstListBox.AddItem wks.name
    Next

End Sub
