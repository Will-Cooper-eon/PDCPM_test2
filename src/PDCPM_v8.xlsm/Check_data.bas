Attribute VB_Name = "Check_data"
Option Explicit

Public k As Long
Public safeToExport As String

Sub check_data() 'Optional ByVal DuplicatesFixed As Boolean = False)

'If no data is present then dont run
If WorksheetFunction.CountA(Range("B4:B5")) = 0 Then
    MsgBox "No data to check"
    
Else
    Dim i As Long
    Dim j As Long
    Dim offset_row As Integer
    
    Dim Error_On_Line As Boolean
    Dim Row_Above As Boolean
    
    ExcelProcesses (False)
    
    k = 0
    i = 1
    Sheet2.Range("Error_Log").Offset(1, 0).Resize(100000, 1).ClearContents
    offset_row = Sheet1.Range("Contract_Data_Start").Cells(1, 1).row
    
    'Calls Further checks as this sorts the data first
    Call Sort_FurtherChecks.Further_checks  'DuplicatesFixed)
    
    ' If DuplicatesFixed = False Then
    Do Until Sheet1.Range("Contract_Data_Start").Offset(i, 0) = ""
        
        Error_On_Line = False
        Row_Above = False
    
        For j = 0 To 410
            If j = 15 Then
                If IsEmpty(Sheet1.Range("Contract_Data_Start").Offset(i, j)) = False And Sheet1.Range("Contract_Data_Start").Offset(i, j) = 0 Then
                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                    & ") : Supplier Threshold set as zero - CEE capped at zero, is this the intention?  Leave blank for no threshold."
                    Error_On_Line = True
                    'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                    Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                    k = k + 1
                End If
            End If
            If j = 16 Then
                If IsEmpty(Sheet1.Range("Contract_Data_Start").Offset(i, j)) = False And Sheet1.Range("Contract_Data_Start").Offset(i, j) = 0 Then
                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                    & ") : Customer Threshold set as zero - CEE capped at zero, is this the intention?  Leave blank for no threshold."
                    Error_On_Line = True
                    'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                    Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                    k = k + 1
                End If
            End If
            If j = 18 Then
                If Sheet1.Range("Contract_Data_Start").Offset(i, j) <> 0.9 Then
                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                    & ") : SE cover level set as " & Sheet1.Range("Contract_Data_Start").Offset(i, j) & " expected 0.9 - please check."
                    Error_On_Line = True
                    'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                    Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                    k = k + 1
                End If
            End If
            If j = 21 Then
                If Sheet1.Range("Contract_Data_Start").Offset(i, j) <> 0.9 Then
                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                    & ") : CEE cover level set as " & Sheet1.Range("Contract_Data_Start").Offset(i, j) & " expected 0.9 - please check."
                    Error_On_Line = True
                    'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                    Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                    k = k + 1
                End If
            End If
            If j = 0 Or j = 2 Or j = 3 Or j = 7 Or j = 8 Or j = 410 Then
                If Application.WorksheetFunction.IsNonText(Sheet1.Range("Contract_Data_Start").Offset(i, j)) = True Then
                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                    & ") : Non String"
                    Error_On_Line = True
                    'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                    Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                    k = k + 1
                End If
            ElseIf j = 409 Then
                If Not (Sheet1.Range("Contract_Data_Start").Offset(i, j) = "Y" Or Sheet1.Range("Contract_Data_Start").Offset(i, j) = "N") Then
                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                    & ") : Renewal flag must be Y or N"
                    Error_On_Line = True
                    'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                    Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                End If
            ElseIf (j > 13 And j < 24) Or (j > 26 And j < 409) Then
                If IsEmpty(Sheet1.Range("Contract_Data_Start").Offset(i, j)) = False And _
                Application.WorksheetFunction.IsNumber(Sheet1.Range("Contract_Data_Start").Offset(i, j)) = False Then
                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                    & ") : Non Blank, Non Number"
                    Error_On_Line = True
                    'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                    Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                    k = k + 1
                End If
            ElseIf j > 1 Then
                If Application.WorksheetFunction.IsNumber(Sheet1.Range("Contract_Data_Start").Offset(i, j)) = False Then
                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                    & ") : Non Number"
                    Error_On_Line = True
                    'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                    Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                    k = k + 1
                Else
                    If j = 5 And Sheet1.Range("Contract_Data_Start").Offset(i, j) < Sheet1.Range("Contract_Data_Start").Offset(i, j - 1) Then
                        Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                        "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                        & ") : Start Date less than Signature Date"
                        Error_On_Line = True
                        'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                        Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                        k = k + 1
    
                    ElseIf j = 6 Then
                        If Sheet1.Range("Contract_Data_Start").Offset(i, j) > DateAdd("m", 120, Sheet1.Range("Contract_Data_Start").Offset(i, j - 1)) - 1 Then
                            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                            "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                            & ") : Contract length is greater than 10 years"
                            Error_On_Line = True
                            'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                            Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                            k = k + 1
                        ElseIf Sheet1.Range("Contract_Data_Start").Offset(i, j) <= Sheet1.Range("Contract_Data_Start").Offset(i, j - 1) Then
                            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                            "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                            & ") : End Date less than or equal to Start Date"
                            Error_On_Line = True
                            'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                            Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                            k = k + 1
                        End If
                    ElseIf j = 4 And Sheet1.Range("Contract_Data_Start").Offset(i, j) > Date Then
                        Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                        "Contract " & i & " column " & j + 2 & " (" & Sheet1.Range("Contract_Data_Start").Offset(0, j) _
                        & ") : Signature Date is in the future"
                        Error_On_Line = True
                        'Range(Cells(i + offset_row, Range("Counter_Party").Column), Cells(i + offset_row, Range("Cust_ID").Column)).Interior.Color = RGB(255, 192, 203)
                        Cells(i + offset_row, j + 2).Interior.ColorIndex = 3
                        k = k + 1
                    End If
                    
                End If
            End If
        Next j
        
        ' highlight entire row
        If Error_On_Line = True Then Call Highlight_Entire_Row(2, 412, i + offset_row, Sheet1, Row_Above)
        
        i = i + 1
        
    Loop
    ' End If
    
    Sheets("Input").Sort.SortFields.Clear 'This solves error with file (Removed Records: Sorting from /xl/worksheets/sheet1.xml part)
    
    MsgBox k & " Errors"
    
    'If no errors were found, and the actual sheet is populated with data then allow user to export errors
    If k = 0 Then
        ThisWorkbook.Sheets("Input").Shapes("Button 12").Visible = False ' export errors
        ThisWorkbook.Sheets("Input").Shapes("Button 19").Visible = False ' Fix Multiple ID contracts
        If ThisWorkbook.Sheets("Input").Range("Modified?") = "N" Then
            ThisWorkbook.Sheets("Input").Shapes("Button 16").Visible = True ' export to CPM
            ThisWorkbook.Sheets("Input").Shapes("Button 18").Visible = True ' Callibration Export
            ThisWorkbook.Sheets("Input").Shapes("Button 21").Visible = False ' Create modified workbook
            MsgBox "Can now export to CPM"
        ElseIf ThisWorkbook.Sheets("Input").Range("Modified?") = "Y" Then
            ThisWorkbook.Sheets("Input").Shapes("Button 16").Visible = False ' export to CPM
            ThisWorkbook.Sheets("Input").Shapes("Button 18").Visible = False ' Callibration Export
            ThisWorkbook.Sheets("Input").Shapes("Button 21").Visible = True ' Create modified workbook
            MsgBox "Can now create new modified portfolio data workbook"
        End If
        safeToExport = "yes"
    ElseIf k > 0 Then
        ' If errors were found, allow them to export the errors
        ThisWorkbook.Sheets("Input").Shapes("Button 12").Visible = True ' export errors
        ThisWorkbook.Sheets("Input").Shapes("Button 19").Visible = True ' Fix Multiple ID contracts
        ThisWorkbook.Sheets("Input").Shapes("Button 16").Visible = False ' export to CPM
        ThisWorkbook.Sheets("Input").Shapes("Button 18").Visible = False ' Callibration Export
        ThisWorkbook.Sheets("Input").Shapes("Button 21").Visible = False ' Create modified workbook
        Sheet2.Select
    End If
End If

    ExcelProcesses (True)

End Sub
