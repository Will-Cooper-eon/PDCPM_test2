Option Explicit

'' The classes EventListener, CustomActions, MyCustomActions are not imported automatically by the build module.
'' After they are imported manually, the comments below can be uncommented. This will enable automatic code exports on save
'' and automatic code imports on open.

'' Clean up our private members
Public Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Print "project thisWorkbook_BeforeClose()"
    SourceControl.deleteMenu
End Sub

Public Sub Workbook_AfterSave(ByVal success As Boolean)
    On Error GoTo App_WorkbookAfterSave_Error

    'Export all the modules for this work book if save was successful
    If success Then
        SourceControl.exportVbaCode ThisWorkbook.VBProject
        SourceControl.exportNamedRanges ThisWorkbook
        MsgBox "Finished saving workbook: " & ThisWorkbook.name & ". Code is exported."
    Else
        MsgBox "Saving workbook: " & ThisWorkbook.name & " was not successful. Code is not exported."
    End If

    Exit Sub
App_WorkbookAfterSave_Error:
    SourceControl.handleError "" & ThisWorkbook.VBProject & ".EventListener afterSave"
End Sub


'Private Sub App_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    If Not Cancel Then
'        SourceControl.formatProject wb.VBProject
'    End If
'End Sub

Public Sub Workbook_Open()
    On Error GoTo App_WorkbookOpen_Error
    
    SourceControl.createMenu

    'Import all the modules for this workbook
    Dim importNow As Integer
    importNow = MsgBox("Import the code for " & ThisWorkbook.name & " now?", vbYesNo, "EventListener Workbook open event")
    If importNow = vbYes Then
        SourceControl.importVbaCode ThisWorkbook.VBProject
        SourceControl.importNamedRanges ThisWorkbook
    End If

    Exit Sub
App_WorkbookOpen_Error:
    SourceControl.handleError "" & ThisWorkbook.VBProject & ".EventListener WorkbookOpen"
End Sub