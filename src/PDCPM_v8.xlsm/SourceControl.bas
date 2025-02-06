Attribute VB_Name = "SourceControl"
Option Explicit


'''''''''''''''''''''''' BUILD ''''''''''''''''''''''''''''''''''''''''''''''
Private Const IMPORT_DELAY As String = "00:00:03"

'We need to make these variables public such that they can be given as arguments to application.ontime()
Public componentsToImport As Dictionary 'Key = componentName, Value = componentFilePath
Public sheetsToImport As Dictionary 'Key = componentName, Value = File object
Public vbaProjectToImport As VBProject

' Returns the directory where code is exported to or imported from.
' When createIfNotExists:=True, the directory will be created if it does not exist yet.
' This is desired when we get the directory for exporting.
' When createIfNotExists:=False and the directory does not exist, an empty String is returned.
' This is desired when we get the directory for importing.
'
' Directory names always end with a '\', unless an empty string is returned.
' Usually called with: fullWorkbookPath = wb.FullName or fullWorkbookPath = vbProject.fileName
' if the workbook is new and has never been saved,
' vbProject.fileName will throw an error while wb.FullName will return a name without slashes.
Public Function getSourceDir(fullWorkbookPath As String, createIfNotExists As Boolean) As String
    ' First check if the fullWorkbookPath contains a \.
    If Not InStr(fullWorkbookPath, "\") > 0 Then
        'In this case it is a new workbook, we skip it
        Exit Function
    End If

    Dim FSO As New Scripting.FileSystemObject
    Dim projDir As String
    projDir = FSO.GetParentFolderName(fullWorkbookPath) & "\"
    Dim srcDir As String
    srcDir = projDir & "src\"
    Dim exportDir As String
    exportDir = srcDir & FSO.GetFileName(fullWorkbookPath) & "\"

    If createIfNotExists Then
        If Not FSO.FolderExists(srcDir) Then
            FSO.CreateFolder srcDir
            Debug.Print "Created Folder " & srcDir
        End If
        If Not FSO.FolderExists(exportDir) Then
            FSO.CreateFolder exportDir
            Debug.Print "Created Folder " & exportDir
        End If
    Else
        If Not FSO.FolderExists(exportDir) Then
            Debug.Print "Folder does not exist: " & exportDir
            exportDir = ""
        End If
    End If
    getSourceDir = exportDir
End Function


' Usually called after the given workbook is saved
Public Sub exportVbaCode(vbaProject As VBProject)
    Dim vbProjectFileName As String
    On Error Resume Next
    'this can throw if the workbook has never been saved.
    vbProjectFileName = vbaProject.fileName
    On Error GoTo 0
    If vbProjectFileName = "" Then
        'In this case it is a new workbook, we skip it
        Debug.Print "No file name for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    Dim export_path As String
    export_path = getSourceDir(vbProjectFileName, createIfNotExists:=True)

    Debug.Print "exporting to " & export_path
    'export all components
    Dim component As VBComponent
    For Each component In vbaProject.VBComponents
        'lblStatus.Caption = "Exporting " & proj_name & "::" & component.Name
        If hasCodeToExport(component) Then
            'Debug.Print "exporting type is " & component.Type
            Select Case component.Type
                Case vbext_ct_ClassModule
                    exportComponent export_path, component
                Case vbext_ct_StdModule
                    exportComponent export_path, component, ".bas"
                Case vbext_ct_MSForm
                    exportComponent export_path, component, ".frm"
                Case vbext_ct_Document
                    exportLines export_path, component
                Case Else
                    'Raise "Unkown component type"
            End Select
        End If
    Next component
End Sub


Private Function hasCodeToExport(component As VBComponent) As Boolean
    hasCodeToExport = True
    If component.CodeModule.CountOfLines <= 2 Then
        Dim firstLine As String
        firstLine = Trim(component.CodeModule.lines(1, 1))
        'Debug.Print firstLine
        hasCodeToExport = Not (firstLine = "" Or firstLine = "Option Explicit")
    End If
End Function


'To export everything else but sheets
Private Sub exportComponent(exportPath As String, component As VBComponent, Optional extension As String = ".cls")
    Debug.Print "exporting " & component.name & extension
    component.Export exportPath & "\" & component.name & extension
End Sub


'To export sheets
Private Sub exportLines(exportPath As String, component As VBComponent)
    Dim extension As String: extension = ".sheet.cls"
    Dim fileName As String
    fileName = exportPath & "\" & component.name & extension
    Debug.Print "exporting " & component.name & extension
    'component.Export exportPath & "\" & component.name & extension
    Dim FSO As New Scripting.FileSystemObject
    Dim outStream As TextStream
    Set outStream = FSO.CreateTextFile(fileName, True, False)
    outStream.Write (component.CodeModule.lines(1, component.CodeModule.CountOfLines))
    outStream.Close
End Sub


' Usually called after the given workbook is opened. The option includeClassFiles is False by default because
' they don't import correctly from VBA. They'll have to be imported manually instead.
Public Sub importVbaCode(vbaProject As VBProject, Optional includeClassFiles As Boolean = False)
    Dim vbProjectFileName As String
    On Error Resume Next
    'this can throw if the workbook has never been saved.
    vbProjectFileName = vbaProject.fileName
    On Error GoTo 0
    If vbProjectFileName = "" Then
        'In this case it is a new workbook, we skip it
        Debug.Print "No file name for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    Dim export_path As String
    export_path = getSourceDir(vbProjectFileName, createIfNotExists:=False)
    If export_path = "" Then
        'The source directory does not exist, code has never been exported for this vbaProject.
        Debug.Print "No import directory for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If

    'initialize globals for Application.OnTime
    Set componentsToImport = New Dictionary
    Set sheetsToImport = New Dictionary
    Set vbaProjectToImport = vbaProject

    Dim FSO As New Scripting.FileSystemObject
    Dim projContents As Folder
    Set projContents = FSO.GetFolder(export_path)
    Dim file As Object
    For Each file In projContents.Files()
        'check if and how to import the file
        checkHowToImport file, includeClassFiles
    Next

    Dim componentName As String
    Dim vComponentName As Variant
    'Remove all the modules and class modules
    For Each vComponentName In componentsToImport.Keys
        componentName = vComponentName
        removeComponent vbaProject, componentName
    Next
    'Then import them
    'Call importComponents
    Debug.Print "Invoking 'importComponents' with Application.Ontime with delay " & IMPORT_DELAY
    ' to prevent duplicate modules, like MyClass1 etc.
    Application.OnTime Now() + TimeValue(IMPORT_DELAY), "'importComponents'"
    Debug.Print "almost finished importing code for " & vbaProject.name
End Sub


Private Sub checkHowToImport(file As Object, includeClassFiles As Boolean)
    Dim fileName As String
    fileName = file.name
    Dim componentName As String
    componentName = Left(fileName, InStr(fileName, ".") - 1)
'    If componentName = "Build" Then
'        '"don't remove or import ourself
'        Exit Sub
'    End If
    If componentName = "SourceControl" Then
        '"don't remove or import ourself
        Exit Sub
    End If

    If Len(fileName) > 4 Then
        Dim lastPart As String
        lastPart = Right(fileName, 4)
        Select Case lastPart
            Case ".cls" ' 10 == Len(".sheet.cls")
                If Len(fileName) > 10 And Right(fileName, 10) = ".sheet.cls" Then
                    'import lines into sheet: importLines vbaProjectToImport, file
                    sheetsToImport.Add componentName, file
                Else
                    ' .cls files don't import correctly because of a bug in excel, therefore we can exclude them.
                    ' In that case they'll have to be imported manually.
                    If includeClassFiles Then
                        'importComponent vbaProject, file
                        componentsToImport.Add componentName, file.Path
                    End If
                End If
            Case ".bas", ".frm"
                'importComponent vbaProject, file
                componentsToImport.Add componentName, file.Path
            Case Else
                'do nothing
                Debug.Print "Skipping file " & fileName
        End Select
    End If
End Sub


' Only removes the vba component if it exists
Private Sub removeComponent(vbaProject As VBProject, componentName As String)
    If componentExists(vbaProject, componentName) Then
        Dim c As VBComponent
        Set c = vbaProject.VBComponents(componentName)
        Debug.Print "removing " & c.name
        vbaProject.VBComponents.Remove c
    End If
End Sub


Public Sub importComponents()
    If componentsToImport Is Nothing Then
        Debug.Print "Failed to import! Dictionary 'componentsToImport' was not initialized."
        Exit Sub
    End If
    Dim componentName As String
    Dim vComponentName As Variant
    For Each vComponentName In componentsToImport.Keys
        componentName = vComponentName
        importComponent vbaProjectToImport, componentsToImport(componentName)
    Next

    'Import the sheets
    For Each vComponentName In sheetsToImport.Keys
        componentName = vComponentName
        importLines vbaProjectToImport, sheetsToImport(componentName)
    Next

    Debug.Print "Finished importing code for " & vbaProjectToImport.name
    'We're done, clear globals explicitly to free memory.
    Set componentsToImport = Nothing
    Set vbaProjectToImport = Nothing
End Sub


' Assumes any component with same name has already been removed.
Private Sub importComponent(vbaProject As VBProject, filePath As String)
    Debug.Print "Importing component from  " & filePath
    'This next line is a bug! It imports all classes as modules!
    vbaProject.VBComponents.Import filePath
End Sub


Private Sub importLines(vbaProject As VBProject, file As Object)
    Dim componentName As String
    componentName = Left(file.name, InStr(file.name, ".") - 1)
    Dim c As VBComponent
    If Not componentExists(vbaProject, componentName) Then
        ' Create a sheet to import this code into. We cannot set the ws.codeName property which is read-only,
        ' instead we set its vbComponent.name which leads to the same result.
        Dim addedSheetCodeName As String
        addedSheetCodeName = addSheetToWorkbook(componentName, vbaProject.fileName)
        Set c = vbaProject.VBComponents(addedSheetCodeName)
        c.name = componentName
    End If
    Set c = vbaProject.VBComponents(componentName)
    Debug.Print "Importing lines from " & componentName & " into component " & c.name

    ' At this point compilation errors may cause a crash, so we ignore those.
    On Error Resume Next
    c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
    c.CodeModule.AddFromFile file.Path
    On Error GoTo 0
End Sub


Public Function componentExists(ByRef proj As VBProject, name As String) As Boolean
    On Error GoTo doesnt
    Dim c As VBComponent
    Set c = proj.VBComponents(name)
    componentExists = True
    Exit Function
doesnt:
    componentExists = False
End Function


' Returns a reference to the workbook. Opens it if it is not already opened.
' Raises error if the file cannot be found.
Public Function openWorkbook(ByVal filePath As String) As Workbook
    Dim wb As Workbook
    Dim fileName As String
    fileName = Dir(filePath)
    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(filePath) 'can raise error
    End If
    Set openWorkbook = wb
End Function


' Returns the CodeName of the added sheet or an empty String if the workbook could not be opened.
Public Function addSheetToWorkbook(sheetName As String, workbookFilePath As String) As String
    Dim wb As Workbook
    On Error Resume Next 'can throw if given path does not exist
    Set wb = openWorkbook(workbookFilePath)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Dim ws As Worksheet
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = sheetName
        'ws.CodeName = sheetName: cannot assign to read only property
        Debug.Print "Sheet added " & sheetName
        addSheetToWorkbook = ws.CodeName
    Else
        Debug.Print "Skipping file " & sheetName & ". Could not open workbook " & workbookFilePath
        addSheetToWorkbook = ""
    End If
End Function


'''''''''''''''''''' ERROR HANDLING ''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RaiseError(errNumber As Integer, Optional errSource As String = "", Optional errDescription As String = "")
    If errSource = "" Then
        'set default values
        errSource = Err.Source
        errDescription = Err.Description
    End If
    Err.Raise vbObjectError + errNumber, errSource, errDescription
End Sub


Public Sub handleError(Optional errLocation As String = "")
    Dim errorMessage As String
    errorMessage = "Error in " & errLocation & ", [" & Err.Source & "] : error number " & Err.Number & vbNewLine & Err.Description
    Debug.Print errorMessage
    MsgBox errorMessage, vbCritical, "vbaDeveloper ErrorHandler"
End Sub


''''''''''''''''''' MENU ''''''''''''''''''''''''''''''''''''''''''''

Public Sub createMenu()
    Dim rootMenu As CommandBarPopup

    'Add the top-level menu to the ribbon Add-ins section
    Set rootMenu = Application.CommandBars(1).Controls.Add(Type:=msoControlPopup, _
    Before:=10, _
    Temporary:=True)
    rootMenu.caption = "vbaDeveloper"

    Dim exSubMenu As CommandBarPopup
    Dim imSubMenu As CommandBarPopup
    'Dim formatSubMenu As CommandBarPopup
    Set exSubMenu = addSubmenu(rootMenu, 1, "Export code for ...")
    Set imSubMenu = addSubmenu(rootMenu, 2, "Import code for ...")
    'Set formatSubMenu = addSubmenu(rootMenu, 3, "Format code for ...")
    addMenuSeparator rootMenu
    Dim refreshItem As CommandBarButton
    Set refreshItem = addMenuItem(rootMenu, "refreshMenu", "Refresh this menu")
    refreshItem.FaceId = 37

    ' menuItem.FaceId = FaceId ' set a picture
    Dim vProject As Variant
    For Each vProject In Application.VBE.VBProjects
        ' We skip over unsaved projects where project.fileName throws error
        On Error GoTo nextProject
        Dim project As VBProject
        Set project = vProject
        Dim projectName As String, caption As String

        projectName = project.name
        caption = projectName & " (" & Dir(project.fileName) & ")" '<- this can throw error

        Dim exCommand As String, imCommand As String, formatCommand As String
        exCommand = "'exportVbProject """ & project.fileName & """'"
        imCommand = "'importVbProject """ & project.fileName & """'"
        ' formatCommand = "'formatVbProject """ & project.fileName & """'"

        addMenuItem exSubMenu, exCommand, caption
        addMenuItem imSubMenu, imCommand, caption
        ' addMenuItem formatSubMenu, formatCommand, caption
nextProject:
    Next vProject
    On Error GoTo 0 'reset the error handling

    'Add menu items for creating and rebuilding XML files
    Dim xmlMenu As CommandBarPopup, exXmlSubMenu As CommandBarPopup
    Set xmlMenu = Application.CommandBars(1).Controls.Add(Type:=msoControlPopup, _
    Before:=10, _
    Temporary:=True)
    xmlMenu.caption = "XML Import-Export"

    Set exXmlSubMenu = addSubmenu(xmlMenu, 1, "Export XML for ...")
    Dim rebuildButton As CommandBarButton
    Set rebuildButton = addMenuItem(xmlMenu, "menu_rebuildXML", "Rebuild a file")
    rebuildButton.FaceId = 35
    Set refreshItem = addMenuItem(xmlMenu, "refreshMenu", "Refresh this menu")
    refreshItem.FaceId = 37

    'add menu items for all open files
    Dim fileName As String
    Dim openFile As Workbook
    For Each openFile In Application.Workbooks
        fileName = openFile.name
        Call addMenuItem(exXmlSubMenu, "'exportXML """ & fileName & """'", fileName)
    Next openFile

End Sub


Private Function addMenuItem(menu As CommandBarPopup, ByVal onAction As String, ByVal caption As String) As CommandBarButton
    Dim menuItem As CommandBarButton
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
    menuItem.onAction = onAction
    menuItem.caption = caption
    Set addMenuItem = menuItem
End Function


Private Function addSubmenu(menu As CommandBarPopup, ByVal position As Integer, ByVal caption As String) As CommandBarPopup
    Dim subMenu As CommandBarPopup
    Set subMenu = menu.Controls.Add(Type:=msoControlPopup)
    subMenu.onAction = position
    subMenu.caption = caption
    Set addSubmenu = subMenu
End Function


Private Sub addMenuSeparator(menuItem As CommandBarPopup)
    menuItem.BeginGroup = True
End Sub


'This sub should be executed when the workbook is closed
Public Sub deleteMenu()
    'For each control, check if its name matches the names of our custom menus - using this method deletes multiple instances of the menu in case duplicates are mistakenly created.
    Dim cbControl
    On Error Resume Next
    For Each cbControl In CommandBars(1).Controls               'TODO if more menus are added, should use a collection instead of multiple if statements (keep code DRY)
        If cbControl.caption = "vbaDeveloper" Then
            Debug.Print "Deleting" & "vbaDeveloper"
            cbControl.Delete
        End If
        If cbControl.caption = "XML Import-Export" Then
            Debug.Print "Deleting" & "XML Import-Export"
            cbControl.Delete
        End If
    Next cbControl
    On Error GoTo 0
End Sub

Public Sub refreshMenu()
    deleteMenu
    createMenu
End Sub

Public Sub exportVbProject(ByVal projectPath As String)
    On Error GoTo exportVbProject_Error

    Dim project As VBProject
    Set project = GetProjectByPath(projectPath)
    exportVbaCode project
    Dim wb As Workbook
    Set wb = openWorkbook(project.fileName)
    exportNamedRanges wb
    MsgBox "Finished exporting code for: " & project.name

    Exit Sub
exportVbProject_Error:
    handleError "exportVbProject"
End Sub


Public Sub importVbProject(ByVal projectPath As String)
    On Error GoTo importVbProject_Error

    Dim project As VBProject
    Set project = GetProjectByPath(projectPath)
    importVbaCode project
    Dim wb As Workbook
    Set wb = openWorkbook(project.fileName)
    importNamedRanges wb
    MsgBox "Finished importing code for: " & project.name

    On Error GoTo 0
    Exit Sub
importVbProject_Error:
    handleError "importVbProject"
End Sub


'Public Sub formatVbProject(ByVal projectPath As String)
'    On Error GoTo formatVbProject_Error
'
'    Dim project As VBProject
'    Set project = GetProjectByPath(projectPath)
'    formatProject project
'    MsgBox "Finished formatting code for: " & project.name & vbNewLine _
'    & vbNewLine _
'    & "Did you know you can also format your code, while writing it, by typing 'application.Run ""formatActiveCodePane""' in the immediate window?"
'
'    On Error GoTo 0
'    Exit Sub
'formatVbProject_Error:
'    handleError "formatVbProject"
'End Sub


Public Sub exportXML(ByVal fileShortName As String)
    'Ask them if they want to save the file first. Warn that existing files could be overwritten. Default to 'Cancel'
    Dim validateChoice As Integer, prompt As String, title As String
    prompt = "Are you sure you want to export " & fileShortName & " to XML? Any previously exported XML data for that file will be overwritten."
    title = "Overwrite existing XML?"
    validateChoice = MsgBox(prompt, vbYesNoCancel, title)

    prompt = "Do you want to save the file before exporting? If unsaved, the exported version will reflect only changes until your most recent save."
    title = "Save file first?"
    validateChoice = MsgBox(prompt, vbYesNoCancel, title)
    If validateChoice = vbCancel Then Exit Sub
    If validateChoice = vbYes Then
        Dim wkb As Workbook
        Set wkb = Workbooks(fileShortName)
        wkb.Save
    End If

    Call unpackXML(fileShortName)
    MsgBox ("File successfully exported to XML. Check the 'src' folder where the file is saved.")
End Sub

Public Sub menu_rebuildXML()
    'This sub lets the user browse to a folder, sets the destination folder as two levels up the folder tree, and then calls the 'rebuildXML' function to zip up the XML data into an Excel file
    Dim destinationFolder As String, containingFolderName As String, errorFlag As Boolean, errorMessage As String
    destinationFolder = "C:\"
    containingFolderName = "C:\"

    containingFolderName = GetFolder(destinationFolder)                                                 'Select containing folder using file picker
    containingFolderName = removeSlash(containingFolderName)                                'Remove trailing slash if it exists

    'destinationFolder is two levels up from the containing folder
    On Error GoTo folderError
    destinationFolder = containingFolderName
    destinationFolder = Left(destinationFolder, Len(destinationFolder) - (Len(destinationFolder) - InStrRev(destinationFolder, "\") + 1)) 'up one level
    destinationFolder = Left(destinationFolder, Len(destinationFolder) - (Len(destinationFolder) - InStrRev(destinationFolder, "\") + 1)) 'up another level
    On Error GoTo 0

    errorFlag = False
    Call rebuildXML(destinationFolder, containingFolderName, errorFlag, errorMessage)

folderError:
    If Err.Number <> 0 Then
        errorFlag = True
        errorMessage = "That's not a valid folder"
    End If

    'Report the status to the user
    If errorFlag = True Then
        MsgBox (errorMessage)
    Else
        MsgBox ("File succesfully rebuilt to here: " & vbCrLf & destinationFolder)
    End If

End Sub

Function GetFolder(InitDir As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    sItem = InitDir
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .title = "Select a Folder"
        .AllowMultiSelect = False
        If Right(sItem, 1) <> "\" Then
            sItem = sItem & "\"
        End If
        .InitialFileName = sItem
        If .Show <> -1 Then
            sItem = InitDir
        Else
            sItem = .SelectedItems(1)
        End If
    End With
    GetFolder = sItem
    Set fldr = Nothing
End Function


Function GetProjectByPath(ByVal projectPath As String) As VBProject
    'Simple search to find project by file path
    Dim project As VBProject
    For Each project In Application.VBE.VBProjects
        On Error GoTo skipone
        If UCase(project.fileName) = UCase(projectPath) Then
            Set GetProjectByPath = project
            Exit Function
        End If
nextprj:
    Next project
    'If not found return nothing
    Exit Function
skipone:
    Resume nextprj
End Function

''''''''' NAMED RANGES ''''''''''''

'Private Enum columns
'    name = 0
'    RefersTo
'    Comments
'End Enum


' Import named ranges from csv file
' Existing ranges with the same identifier will be replaced.
Public Sub importNamedRanges(wb As Workbook)
    Dim importDir As String
    importDir = getSourceDir(wb.FullName, createIfNotExists:=False)
    If importDir = "" Then
        Debug.Print "No import directory for workbook " & wb.name & ", skipping"
        Exit Sub
    End If

    Dim fileName As String
    fileName = importDir & "NamedRanges.csv"
    Dim FSO As New Scripting.FileSystemObject
    If FSO.FileExists(fileName) Then
        Dim inStream As TextStream
        Set inStream = FSO.OpenTextFile(fileName, ForReading, Create:=False)
        Dim line As String
        Do Until inStream.AtEndOfStream
            line = inStream.ReadLine
            importName wb, line
        Loop
        inStream.Close
    End If
End Sub


Private Sub importName(wb As Workbook, line As String)
    Dim parts As Variant
    parts = Split(line, ",")
    Dim rangeName As String, rangeAddress As String, comment As String
    rangeName = parts(0)
    rangeAddress = parts(1)
    comment = parts(2)

    ' Existing namedRanges don't need to be removed first.
    ' wb.Names.Add will automatically replace or add the given namedRange.
    wb.Names.Add(rangeName, rangeAddress).comment = comment
End Sub


'Export named ranges to csv file
Public Sub exportNamedRanges(wb As Workbook)
    Dim exportDir As String
    exportDir = getSourceDir(wb.FullName, createIfNotExists:=True)
    Dim fileName As String
    fileName = exportDir & "NamedRanges.csv"

    Dim lines As Collection
    Set lines = New Collection
    Dim aName As name
    Dim t As Variant
    For Each t In wb.Names
        Set aName = t
        If hasValidRange(aName) Then
            lines.Add aName.name & "," & aName.RefersTo & "," & aName.comment
        End If
    Next
    If lines.Count > 0 Then
        'We have some names to export
        Debug.Print "writing to  " & fileName

        Dim FSO As New Scripting.FileSystemObject
        Dim outStream As TextStream
        Set outStream = FSO.CreateTextFile(fileName, overwrite:=True, Unicode:=False)
        On Error GoTo closeStream
        Dim line As Variant
        For Each line In lines
            outStream.WriteLine line
        Next line
closeStream:
        outStream.Close
    End If
End Sub


Private Function hasValidRange(aName As name) As Boolean
    On Error GoTo no
    hasValidRange = False
    Dim aRange As Range
    Set aRange = aName.RefersToRange
    hasValidRange = True
no:
End Function


' Clean up all named ranges that don't refer to a valid range.
' This sub is not used by the import and export functions.
' It is provided only for convenience and can be run manually.
Public Sub removeInvalidNamedRanges(wb As Workbook)
    Dim aName As name
    Dim t As Variant
    For Each t In wb.Names
        Set aName = t
        If Not hasValidRange(aName) Then
            aName.Delete
        End If
    Next
End Sub

''''''''''''' XML EXPORTED ''''''''''''''''''''''''''''''
'Public Const XML_FOLDER_NAME = "XMLsource\"
'Public Const TEMP_ZIP_NAME = "temp.zip"

Sub test_unpackXML()
    Call unpackXML("tempDevFile.xlsm")
    MsgBox ("Done")
End Sub

Public Sub unpackXML(fileShortName As String)
    'This unpacks the most recently saved version of the file that is passed as an argument.
    'It's necessary for the file to be currently open; calling function should (if appropriate) ask the user if they want to save before executing so that the version on the hard drive is the most recent.

    Dim fileName As String, exportPath As String, exportPathXML As String
    fileName = Workbooks(fileShortName).FullName
    exportPath = getSourceDir(fileName, createIfNotExists:=True)
    exportPathXML = exportPath & "XMLsource\"

    Dim FSO As New Scripting.FileSystemObject
    If Not FSO.FolderExists(exportPathXML) Then
        FSO.CreateFolder exportPathXML
        Debug.Print "Created Folder " & exportPathXML
    End If

    'Copy file to temp zip file
    Dim tempZipFileName As String
    tempZipFileName = exportPath & "temp.zip"
    'FileCopy fileName, tempZipFileName
    FSO.CopyFile fileName, tempZipFileName, True

    'unzip the temp zip file to the folder
    Call Unzip(tempZipFileName, exportPathXML)

    'delete the temp zip file
    Kill tempZipFileName

End Sub

Sub Unzip(Fname As Variant, DefPath As String)
    'Code modified from example found here: http://www.rondebruin.nl/win/s7/win002.htm
    Dim FSO As Object
    Dim oApp As Object
    Dim FileNameFolder As Variant

    If Fname = False Then
        'Do nothing
    Else
        DefPath = addSlash(DefPath)
        FileNameFolder = DefPath

        'Delete all the files in the folder DefPath first if you want
        On Error Resume Next
        Clear_All_Files_And_SubFolders_In_Folder (DefPath)
        On Error GoTo 0

        'Extract the files into the Destination folder
        Set oApp = CreateObject("Shell.Application")
        oApp.Namespace("" & FileNameFolder).CopyHere oApp.Namespace("" & Fname).Items 'The ""&  is to address a bug - for some reason VBA doesn't like to use the passed strings in this situation. Found discussion on this here: http://forums.codeguru.com/showthread.php?443782-CreateObject(-quot-Shell-Application-quot-)-Error

        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
    End If
End Sub


Sub Clear_All_Files_And_SubFolders_In_Folder(MyPath As String)
    'Delete all files and subfolders
    'Be sure that no file is open in the folder
    If Right(MyPath, 1) = "\" Then
        MyPath = Left(MyPath, Len(MyPath) - 1)
    End If

    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")

    If FSO.FolderExists(MyPath) = False Then
        MsgBox MyPath & " doesn't exist"
        Exit Sub
    End If

    On Error Resume Next
    'Delete files
    FSO.DeleteFile MyPath & "\*.*", True
    'Delete subfolders
    FSO.DeleteFolder MyPath & "\*.*", True
    On Error GoTo 0

End Sub

Public Sub rebuildXML(destinationFolder As String, containingFolderName As String, errorFlag As Boolean, errorMessage As String)

    'input format cleanup - containing folder name should not have trailing "\"
    containingFolderName = removeSlash(containingFolderName)
    destinationFolder = removeSlash(destinationFolder)

    'Make sure that the containingFolderName has an XML subfolder
    Dim xmlFolderName As String
    xmlFolderName = containingFolderName & "\" & "XMLsource\"
    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")
    If FSO.FolderExists(xmlFolderName) = False Then
        errorMessage = "We couldn't find XML data in that folder. Make sure you pick the folder under /src that is named the same as the Excel to be rebuilt, and that it contains XML data."
        errorFlag = True
        Exit Sub
    End If

    'Set what some items should be named
    Dim fileExtension As String, strDate As String, fileShortName As String, fileName As String, zipFileName As String
    strDate = VBA.Format(Now, " yyyy-mm-dd hh-mm-ss")
    fileExtension = "." & Right(containingFolderName, Len(containingFolderName) - InStrRev(containingFolderName, "."))  'The containing folder is the folder that is under \src and that is named the same thing as the target file (folder is filename.xlsx) - can parse file ending out of folder
    fileShortName = Right(containingFolderName, Len(containingFolderName) - InStrRev(containingFolderName, "\"))        'This should be just the final folder name
    fileShortName = Left(fileShortName, Len(fileShortName) - (Len(fileShortName) - InStr(fileShortName, ".")) - 1)                            'remove the extension, since we've saved that separately.
    fileName = destinationFolder & "\" & fileShortName & "-rebuilt" & strDate & fileExtension

    zipFileName = containingFolderName & "\" & "temp.zip"

    'Make sure we're not accidentally overwriting anything - this should be rare
    If FSO.FileExists(zipFileName) Then
        errorMessage = "There is already a file named " & "temp.zip" & " in the folder " & containingFolderName & ". This file needs to be removed before continuing."
        errorFlag = True
        Exit Sub
    End If

    'Zip the folder into the FileNameZip
    Call Zip_All_Files_in_Folder(xmlFolderName, zipFileName)

    'Rename the zipFileName to be the fileName (this effectively removes the zip file)
    Name zipFileName As fileName
    errorFlag = False

End Sub



Sub Zip_All_Files_in_Folder(FolderName As Variant, FileNameZip As Variant)
    'Code modified from example found here: http://www.rondebruin.nl/win/s7/win001.htm
    Dim strDate As String, DefPath As String
    Dim oApp As Object

    'Create empty Zip File
    NewZip (FileNameZip)

    Set oApp = CreateObject("Shell.Application")
    'Copy the files to the compressed folder
    oApp.Namespace("" & FileNameZip).CopyHere oApp.Namespace("" & FolderName).Items             '""& added due to bug in VBA

    'Keep script waiting until Compressing is done
    On Error Resume Next
    Do Until oApp.Namespace("" & FileNameZip).Items.Count = _
        oApp.Namespace("" & FolderName).Items.Count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0
End Sub

Sub NewZip(sPath)
    'Create empty Zip File
    'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Function removeSlash(strFolder) As String
    If Right(strFolder, 1) = "\" Then
        strFolder = Left(strFolder, Len(strFolder) - 1)
    End If
    removeSlash = strFolder
End Function
Function addSlash(strFolder) As String
    If Right(strFolder, 1) <> "\" Then
        strFolder = strFolder & "\"
    End If
    addSlash = strFolder
End Function
