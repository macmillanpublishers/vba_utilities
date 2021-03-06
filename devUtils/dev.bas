Attribute VB_Name = "dev"
Option Explicit
Option Base 1


' by Erica Warren - erica.warren@macmillan.com
' Good advice from here: http://www.cpearson.com/excel/vbe.aspx

' ====== USE ==================================================
' For help with VBA development
' Exports all modules in open templates to the local Word-template git repo (ABOVE!!)
' Shared modules go into "SharedMacros" directory, the rest are
' saved in the same directory as the template they live in

' Also imports all modules saved in git repo to the open templates.
' great for dealing with template file merge conflicts!

' ===== DEPENDENCIES ==========================================
' Obviously clone the git repo and add its path ABOVE

' Each template gets its own subdirectory in the repo, name matches exactly (w/o extension)
' Modules that are shared among all template must have name start with "Shared"
' modules that need to be imported into templates but not tracked in git
' are in word-template/dependencies

' Not tested on Mac, because saving templates on Mac causes all kinds of nonsense

' ====== WARNING ==============================================
' advice from http://www.cpearson.com/excel/vbe.aspx
' Many VBA-based computer viruses propagate themselves by creating and/or modifying
' VBA code. Therefore, many virus scanners may automatically and without warning or
' confirmation delete modules that reference the VBProject object, causing a permanent
' and irretrievable loss of code. Consult the documentation for your anti-virus
' software for details.
'
' So be sure to export and commit often!

Sub ExportAllModules()
    ' Exports all VBA modules in all open templates to local git repo
    
    ' Cycle through each open document
    Dim oDoc As Document
    Dim strExtension As String
    Dim oProject As VBIDE.VBProject
    Dim oModule As VBIDE.VBComponent
    Dim strTemplateModules As String
    Dim strDependencies As String
    Dim strDepFiles As String
    Dim strEachFile As String
    Dim strRepoPath As String
    Dim openTemplates As Collection
    Set openTemplates = New Collection
    
    For Each oDoc In Documents
        Debug.Print oDoc.Name
        ' Separate the name and the extension of the document
        strExtension = Right(oDoc.Name, Len(oDoc.Name) - _
            (InStrRev(oDoc.Name, ".") - 1))
        
        ' We just want to work with .dotm and .docm (others can't have macros)
        If strExtension = ".dotm" Or strExtension = ".docm" Then
            ' later need to close > copy > open these files, but if we loop
            ' thru Documents collection, "open" will add file back try again.
            ' So create Collection to loop once:
            openTemplates.Add oDoc
            
            ' get FULL path to this template in its repo
            strRepoPath = GetRepoPath(oDoc)
    
            If oDoc.Name = "genUtils.dotm" Then
            ' Modules that need to be imported into templates but that we do
            ' not want to track. We don't want to export these, so let's get
            ' then into a string check against later.
                
                strDependencies = strRepoPath & Application.PathSeparator & _
                    "dependencies"
                Debug.Print strDependencies
                ' Dir() w/ arguments returns first file name that matches
                ' !!!! When switch to submodules, will need to change this to
                ' search subdirectories.
                strEachFile = Dir(strDependencies & Application.PathSeparator & _
                     "*.*", vbNormal)
                Do While Len(strEachFile) > 0
                    Debug.Print strEachFile
                    strDepFiles = strDepFiles & strEachFile & vbNewLine
'                    Debug.Print strDepFiles
                    ' Dir() again w/o arguments returns the NEXT file that matches orig arguments
                    ' if nothing else matches, returns empty string
                    strEachFile = Dir
                Loop
            Else
                strDepFiles = vbNullString
            End If

            ' Make sure we're referencing the correct project
            Set oProject = oDoc.VBProject
        
            strTemplateModules = strRepoPath & Application.PathSeparator
            
            ' Cycle through each module
            For Each oModule In oProject.VBComponents
                ' Skip modules in dependencies directory
                If InStr(strDepFiles, oModule.Name) = 0 Then
                    ' Don't export forms, they are always wonky. Will have to
                    ' manage manually
                    If oModule.Type <> vbext_ct_MSForm Then
                        Call ExportVBComponent(VBComp:=oModule, _
                            FolderName:=strTemplateModules)
                    End If
                End If
            Next
        End If
    Next oDoc
    
    ' Have to do this in a separate loop if we're opening the files after,
    ' otherwise the newly opened file is added back to the Documents
    ' collection and it keeps looping through them.
'    Dim A As Long
    Dim aDoc As Document
    If openTemplates.Count > 0 Then
'        For A = 1 To openTemplates.Count
'            Set aDoc = openTemplates.Item(A)
        For Each aDoc In openTemplates
            ' And also save the template file in the repo if it's not open from there
            ' CopyTemplateToRepo closes and re-opens the doc, so don't use it for THIS doc
            If aDoc.Name <> ThisDocument.Name Then
                CopyTemplateToRepo TemplateDoc:=aDoc, OpenAfter:=True
            Else
                'Debug.Print ThisDocument.Name
                 aDoc.Save
            End If
        Next aDoc
    End If


End Sub


Private Sub ExportVBComponent(VBComp As VBIDE.VBComponent, _
                FolderName As String, _
                Optional FileName As String, _
                Optional OverwriteExisting As Boolean = True)

    Dim Extension As String
    Dim FName As String
    
    
    Extension = GetFileExtension(VBComp:=VBComp)
    ' Don't auto-export UserForms, because they often add or remove a single
    ' blank like that gets tracked in git in the code module AND the binary
    ' .frx file. Will have to manage userforms manually
    If Extension <> ".frm" Then
        ' Build full file name of module
        If Trim(FileName) = vbNullString Then
            FName = VBComp.Name & Extension
        Else
            FName = FileName
            If InStr(1, FName, ".", vbBinaryCompare) = 0 Then
                FName = FName & Extension
            End If
        End If
        
        ' Can't delete ThisDocument.cls module, but doesn't always have code
        ' So don't export if empty
        If VBComp.CodeModule.CountOfLines <> 0 Then
        
            ' Build full path to save module to
            If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
                FName = FolderName & FName
            Else
                FName = FolderName & "\" & FName
            End If
        
    
            ' delete previous version of module
            If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
                If OverwriteExisting = True Then
                    Kill FName
                Else
                    Exit Sub
                End If
            End If
    
            ' Export the module
            VBComp.Export FileName:=FName
        End If
    End If
    'Debug.Print FName
    
    ' ======================================
    ' Was attempting to checkout UserForm binary after export, since git almost
    ' always tracked modifications even when none are made, but it wasn't
    ' quite working so we'll just skip it (see above)
'    If Extension = ".frm" Then
'        Dim strBinaryFile As String
'
'        strBinaryFile = Left(FName, Len(FName) - 1) & "x"
'        'Debug.Print strBinaryFile
'
'        Dim strShellCmd As String
'        strShellCmd = "cmd.exe /C C: & cd " & strRepoPath & " & git checkout " & strBinaryFile
'        strShellCmd = Replace(strShellCmd, "\", "\\")
'
'        'Debug.Print strShellCmd
'
'        Dim result As Variant
'
'        result = Shell(strShellCmd, vbMinimizedNoFocus)
'        'Debug.Print result
'    End If
    
    End Sub
    
Private Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' COPIED FROM http://www.cpearson.com/excel/vbe.aspx
' This returns the appropriate file extension based on the Type of
' the VBComponent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case VBComp.Type
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case vbext_ct_Document
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ".bas"
    End Select
    
End Function


Sub ImportAllModules()
  ' Removes all modules in all open template
  ' and reimports them from the local Word-template git repo
  ' SO BE SURE THE MODULES IN THE REPO ARE UP TO DATE
  
  Dim oDocument As Document
  Dim strExtension As String              ' extension of current document
  Dim strSubDirName As String             ' name of subdirectory of template in repo
  Dim strDirInRepo() As String      ' declare number of items in array
  Dim strModuleExt(1 To 3) As String     ' declare number of items in array
  Dim strModuleFileName As String         ' file name with extension, no path
  Dim A As Long
  Dim B As Long
  Dim Counter As Long
  Dim VBComp As VBIDE.VBComponent     ' object for module we're importing
  Dim strFullModulePath As String     ' full path to module with extension
  Dim strModuleName As String         ' Just the module name w/ no extension
  Dim tempVBComp As VBIDE.VBComponent ' Temp module to import ThisDocument code
  Dim currentVBProject As VBIDE.VBProject     ' object of the VB project the modules are in
  Dim strNewCode As String            ' New code in ThisDocument.cls module
  Dim strDependencies As String
  Dim strEachFile As String
  Dim openTemplates As Collection
  Set openTemplates = New Collection
  
  For Each oDocument In Documents
    Debug.Print oDocument.Name
      
    ' We don't want to run this on this code here
    If oDocument.Name <> ThisDocument.Name Then
      ' Separate the name and the extension of the document
      strExtension = Right(oDocument.Name, Len(oDocument.Name) - _
          (InStrRev(oDocument.Name, ".") - 1))
      strSubDirName = Left(oDocument.Name, InStrRev(oDocument.Name, ".") - 1)
      'Debug.Print "File name is " & oDocument.Name
      'Debug.Print "Extension is " & strExtension
      'Debug.Print "Directory is " & strSubDirName
      
      ' We just want to work with .dotm and .docm (others can't have macros)
      If strExtension = ".dotm" Or strExtension = ".docm" Then
        ' later need to close > copy > open these files, but if we loop
        ' thru Documents collection, "open" will add file back try again.
        ' So create Collection to loop once:
        openTemplates.Add oDocument
        
        ' get FULL path to this template in its repo
        ReDim strDirInRepo(1 To 1)
        strDirInRepo(1) = GetRepoPath(oDocument)

        If oDocument.Name = "genUtils.dotm" Then
        ' Modules that need to be imported into templates but that we do
        ' not want to track. We do want to import these, so let's get
        ' them into a string check against later.
          
          strDependencies = strDirInRepo(1) & Application.PathSeparator & _
              "dependencies"
          Debug.Print strDependencies
          ReDim Preserve strDirInRepo(1 To 2)
          strDirInRepo(2) = strDependencies
        End If
                      
        ' an array of file extensions we're importing, since there are other files in the repo
        strModuleExt(1) = "bas"
        strModuleExt(2) = "cls"
        strModuleExt(3) = "frm"
        
        ' Get rid of all code currently in there, so we don't create duplicates
        Call DeleteAllVBACode(oDocument)
        
        ' set the Project object for this document
        Set currentVBProject = Nothing
        Set currentVBProject = oDocument.VBProject
        
        ' loop through the directories
        For A = LBound(strDirInRepo()) To UBound(strDirInRepo())
          ' for each directory, loop through all files of each extension
          For B = LBound(strModuleExt()) To UBound(strModuleExt())
            ' Dir function returns first file that matches in that dir
            strModuleFileName = Dir(strDirInRepo(A) & "*." & strModuleExt(B))
            ' so loop through each file of that extension in that directory
            Do While strModuleFileName <> "" And Counter < 100
              Counter = Counter + 1               ' to prevent infinite loops
              'Debug.Print strModuleFileName
              
              strModuleName = Left(strModuleFileName, InStrRev(strModuleFileName, ".") - 1)
              strFullModulePath = strDirInRepo(A) & Application.PathSeparator & strModuleFileName
              'Debug.Print "Full path to module is " & strFullModulePath
              
              ' Resume Next because Set VBComp = current project will cause an error if that
              ' module doesn't exist, and it doesn't because we just deleted everything
              On Error Resume Next
              Set VBComp = Nothing
              Set VBComp = currentVBProject.VBComponents(strModuleName)
              
              ' So if that Set VBComp failed because it doesnt' exist, add it!
              If VBComp Is Nothing Then
                currentVBProject.VBComponents.Import FileName:=strFullModulePath
                Debug.Print strFullModulePath
              Else    ' it DOES exist already
                ' See then if it's the "ThisDocument" module, which can't be deleted
                ' So we can't import because it would just create a duplicate, not replace
                If VBComp.Type = vbext_ct_Document Then
                  ' sp we'll create a temp module of the module we want to import
                  Set tempVBComp = currentVBProject.VBComponents.Import(strFullModulePath)
                  ' then delete the content of ThisDocument and replace it with the content
                  ' of the temp module
                  With VBComp.CodeModule
                      .DeleteLines 1, .CountOfLines
                      strNewCode = tempVBComp.CodeModule.lines(1, tempVBComp.CodeModule.CountOfLines)
                      .InsertLines 1, strNewCode
                  End With
                  On Error GoTo 0
                  ' then remove the temp module
                  currentVBProject.VBComponents.Remove tempVBComp
                End If
              End If
              
              ' calling Dir function again w/ no arguments gets NEXT file that
              ' matches original call. If no more files, returns empty string.
              strModuleFileName = Dir()
            Loop
                        
            'Debug.Print strModuleFileName
          Next B
        Next A
      End If
    End If
  Next oDocument
  ' Have to do this in a separate loop if we're opening the files after,
  ' otherwise the newly opened file is added back to the Documents
  ' collection and it keeps looping through them.
'    Dim A As Long
  Dim aDoc As Document
  If openTemplates.Count > 0 Then
'        For A = 1 To openTemplates.Count
'            Set aDoc = openTemplates.Item(A)
    For Each aDoc In openTemplates
          ' And also save the template file in the repo if it's not open from there
          ' CopyTemplateToRepo closes and re-opens the doc, so don't use it for THIS doc
      If aDoc.Name <> ThisDocument.Name Then
        CopyTemplateToRepo TemplateDoc:=aDoc, OpenAfter:=True
      Else
              'Debug.Print ThisDocument.Name
        aDoc.Save
      End If
    Next aDoc
  End If
    
End Sub


Sub DeleteAllVBACode(objTemplate As Document)
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = objTemplate.VBProject
    
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            Set CodeMod = VBComp.CodeModule
            With CodeMod
                .DeleteLines 1, .CountOfLines
            End With
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Sub


Sub CopyTemplateToRepo(TemplateDoc As Document, Optional OpenAfter As _
    Boolean = True)
' copies the current template file to the local git repo

    Dim strRepoPath As String
    strRepoPath = GetRepoPath(TemplateDoc)
    If strRepoPath <> TemplateDoc.Path Then
        If TemplateDoc.Name <> ThisDocument.Name Then
            Dim strCurrentTemplatePath As String
            Dim strDestinationFilePath As String
            
            ' Current file full path, to use for FileCopy later
            strCurrentTemplatePath = TemplateDoc.FullName
            Debug.Print strCurrentTemplatePath
            
            ' location in repo
            strDestinationFilePath = strRepoPath & Application.PathSeparator & _
                TemplateDoc.Name
            Debug.Print strDestinationFilePath
            
            ' Check if the file is there already
            Dim blnInstalled As Boolean
            blnInstalled = False
            If genUtils.IsInstalledAddIn(TemplateDoc.Name) = True Then
                blnInstalled = True
                AddIns(TemplateDoc.Name).Installed = False
            End If
    
            ' Template needs to be closed for FileCopy to work
            ' ALSO: changing doc properties does NOT count as a "change", so Word
            ' sees the file as unchanged and doesn't actually save, and also
            ' doesn't throw an error so we set Saved = False before saving to get
            ' it working right.
            TemplateDoc.Saved = False
            TemplateDoc.Close SaveChanges:=wdSaveChanges
            Set TemplateDoc = Nothing

            ' copy copy copy copy
            ' but NOT if it's genUtils -- this current file right here has a
            ' reference to it, so we can never copy it ever haha!
            On Error GoTo StupidError
            If strCurrentTemplatePath <> strDestinationFilePath Then
                VBA.FileCopy Source:=strCurrentTemplatePath, _
                    Destination:=strDestinationFilePath
            End If
            On Error GoTo 0

            ' Reinstall add-in if it's a global template
            If blnInstalled = True Then
                WordBasic.DisableAutoMacros     ' Not sure this really works tho
                AddIns(strCurrentTemplatePath).Installed = True
            End If
            
            ' And then open the document again if you wanna.
            ' Though note that AutoExec and Document_Open subs will run when
            ' you do!
            If OpenAfter = True Then
                Documents.Open FileName:=strCurrentTemplatePath, _
                            ReadOnly:=False, _
                            Revert:=False
            End If
        End If
    End If
    Exit Sub
StupidError:
    If Err.Number = 70 And InStr(strCurrentTemplatePath, "genUtils.dotm") > 0 Then
        Resume Next
    End If
End Sub

Sub CheckChangeVersion()
' Display userform with template names and version numbers,
' allow user to enter updated version numbers
' and update the template and version file

' ####### DEPENDENCIES ######
' VersionForm userform module and SharedMacros standard module
    
    
    ' A is for looping through all templates
    Dim A As Long
    Dim lngLBound As Long
    
    ' ===== get array of templates paths ====================
    Dim strFullPathToFinalTemplates() As String
    strFullPathToFinalTemplates = GetTemplatesList(TemplatesYouWant:=allTemplates, PathToRepo:=strRepoPath)
    
    lngLBound = LBound(strFullPathToFinalTemplates)
'    Debug.Print lngLBound
    
    ' ===== build full path to version text file / read current version number file ============
    Dim strFullPathToTextFile() As String
    Dim strCurrentVersion() As String     ' String because can have multiple dots
    
    For A = LBound(strFullPathToFinalTemplates) To UBound(strFullPathToFinalTemplates)
        ReDim Preserve strFullPathToTextFile(lngLBound To A)
        strFullPathToTextFile(A) = LocalPathToRepoPath(LocalPath:=strFullPathToFinalTemplates(A), VersionFile:=True)
'        Debug.Print strFullPathToTextFile(A)
        ReDim Preserve strCurrentVersion(lngLBound To A)
        strCurrentVersion(A) = ReadTextFile(Path:=strFullPathToTextFile(A), FirstLineOnly:=False)
        Debug.Print "Text file in repo : |" & strCurrentVersion(A) & "|"
    Next A
    
    ' ===== get just template name ==========================
    Dim strFileName() As String
    
    For A = LBound(strFullPathToFinalTemplates) To UBound(strFullPathToFinalTemplates)
        ReDim Preserve strFileName(lngLBound To A)
        strFileName(A) = Right(strFullPathToFinalTemplates(A), (InStr(StrReverse(strFullPathToFinalTemplates(A)), _
            Application.PathSeparator)) - 1)
'        Debug.Print strFileName(A)
    Next A
    
    ' ======= create instance of userform, populate with template names/versions ====
    Dim objVersionForm As VersionForm
    Set objVersionForm = New VersionForm

    For A = LBound(strCurrentVersion) To UBound(strCurrentVersion)
        objVersionForm.PopulateFormData A, strFileName(A), strCurrentVersion(A)
    Next A
    
    
    ' ===== display the userform! ===========================
    ' User enters new values, end if they click cancel
    objVersionForm.Show
    
    If objVersionForm.CancelMe = True Then
        Unload objVersionForm
        Exit Sub
    End If
    
    ' ===== check if new versions entered, if so load into array too ====
    Dim strNewVersion() As String
    Dim lngIndexToUpdate() As Long
    Dim B As Long
    
    ' Subtract 1 here so we can add 1 when building array and start at same index
    B = lngLBound - 1
    
    For A = LBound(strCurrentVersion) To UBound(strCurrentVersion)
        ' get new version from userform
        ReDim Preserve strNewVersion(lngLBound To A)
        strNewVersion(A) = objVersionForm.NewVersion(FrameName:=strFileName(A))
'        Debug.Print "New " & A & ": |" & strNewVersion(A) & "|"
        
        ' only update if value is not null and not equal current version number
        If strNewVersion(A) <> vbNullString And strNewVersion(A) <> strCurrentVersion(A) Then
            B = B + 1
            ReDim Preserve lngIndexToUpdate(lngLBound To B)
            
            ' an array of index numbers of the other arrays
            lngIndexToUpdate(B) = A
'            Debug.Print "Update: " & strFileName(lngIndexToUpdate(B))
        End If

    Next A
    
    
    ' ===== if new versions, update files =====
    ' Is anything in our new array?
    
    If B = lngLBound - 1 Then
        Unload objVersionForm
        Exit Sub
    Else
        Dim objTemplateDoc As Document
        
        For B = LBound(lngIndexToUpdate) To UBound(lngIndexToUpdate)
            ' FUTURE:   make sure not on master
            '           make sure working dir is clean?
            '           eventually git stash first, then commit changes (incl templates), then unstash
            
            ' Overwrite text version file in repo with new version number
            OverwriteTextFile TextFile:=strFullPathToTextFile(lngIndexToUpdate(B)), NewText:=strNewVersion(lngIndexToUpdate(B))
            
            ' Open local template file
            Documents.Open FileName:=strFullPathToFinalTemplates(lngIndexToUpdate(B)), ReadOnly:=False, Visible:=False
            Set objTemplateDoc = Nothing
            Set objTemplateDoc = Documents(strFullPathToFinalTemplates(lngIndexToUpdate(B)))
            
            ' Change custom properties to new version number
            objTemplateDoc.CustomDocumentProperties("Version").Value = strNewVersion(lngIndexToUpdate(B))
            
            ' Copy file to repo (it saves and closes the file too)
            CopyTemplateToRepo TemplateDoc:=objTemplateDoc, OpenAfter:=False
            
            Set objTemplateDoc = Nothing
        Next B
    End If
        
    ' ===== maybe also add and commit changes? stash first then unstash at end? =====
    
    Unload objVersionForm
    
    
End Sub


' ===== GetRepoPath ===========================================================
' returns the directory of the git repo for this template file. Saved in a
' custom doc property. If property doesn't exist or is wrong, it will prompt you
' to add the correct path. Obviously will cause some issues if other people are
' updating and pushing the files, but we'll cross that bridge when it happens.

Private Function GetRepoPath(objDoc As Document) As String
    Dim strRepo As String
    On Error GoTo repoError
    strRepo = objDoc.CustomDocumentProperties("repo")
    
    If genUtils.GeneralHelpers.IsItThere(strRepo) = False Then
        Err.Raise 5
    End If
    GetRepoPath = strRepo
    Exit Function
repoError:
    If Err.Number = 5 Then      ' "Invalid procedure call or argument" si.e. prop doesn't exist
        strRepo = InputBox("Enter the full path to the repo for " & objDoc.Name)
        If strRepo <> vbNullString Then
            ' trailing separator if includedd
            If Right(strRepo, 1) = Application.PathSeparator Then
                strRepo = Left(strRepo, Len(strRepo) - 1)
            End If
            ' set do prop for next time
            objDoc.CustomDocumentProperties.Add Name:="repo", LinkToContent:=False, _
                Value:=strRepo, Type:=msoPropertyTypeString
            Resume Next
        Else
            MsgBox "That's not a full path :("
            Exit Function
        End If
    Else
        MsgBox Err.Number & ": " & Err.Description
        Exit Function
    End If
End Function

