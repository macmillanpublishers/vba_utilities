Attribute VB_Name = "SharedFileInstaller"
' created by Erica Warren - erica.warren@macmillan.com

' ======== PURPOSE ===================================
' Downloads and installs an array of template files & logs the downloads

' ======== DEPENDENCIES =======================================
' This is Part 2 of 2. Must be called from a sub in another module that declares file names and locations.
' The template file needs to be uploaded as an attachment to https://confluence.macmillan.com/display/PBL/Test
' If this is an installer file, The Part 1 code needs to reside in the ThisDocument module as a sub called
' Documents_Open in a .docm file so that it will launch when users open the file.
' Requires the SharedMacros module be installed in the same template.

Option Explicit
Option Base 1

Sub Installer(DownloadFrom As GitBranch, Installer As Boolean, TemplateName As String, ByRef TemplatesToInstall() As String)

'"Installer" argument = True if this is for a standalone installtion file.
'"Installer" argument = False is this is part of a daily check of the current file and only updates if out of date.
    
    ' Separate file name from directory path
    Dim lngBreak As Long
    Dim FileName() As String
    Dim FinalDir() As String
    Dim Z As Long
    
    For Z = LBound(TemplatesToInstall()) To UBound(TemplatesToInstall())
'        Debug.Print "Path: " & TemplatesToInstall(z)
        
        lngBreak = InStrRev(TemplatesToInstall(Z), Application.PathSeparator)
'        Debug.Print "Final sep at: " & lngBreak
        
        If lngBreak >= 1 Then
            ReDim Preserve FileName(1 To Z)
            FileName(Z) = Right(TemplatesToInstall(Z), (Len(TemplatesToInstall(Z)) - lngBreak))
            
            ReDim Preserve FinalDir(1 To Z)
            FinalDir(Z) = Left(TemplatesToInstall(Z), lngBreak - 1)
            
'            Debug.Print "File Name #" & z & ": " & FileName(z)
'            Debug.Print "Directory #" & z & ": " & FinalDir(z)
        Else
            ' No path separator in full path specification
            MsgBox "You need to specify the full path to your templates!"
            Exit Sub
        End If
        
    Next Z
    
    
    '' --------------- Set up variable names ----------------------------------------------
    '' Create style directory and logfile names
    Dim A As Long
    Dim arrLogInfo() As Variant
    ReDim arrLogInfo(1 To 3)
    Dim strStyleDir() As String
    ReDim strStyleDir(LBound(FileName()) To UBound(FileName()))
    Dim strLogDir() As String
    ReDim strLogDir(LBound(FileName()) To UBound(FileName()))
    Dim strFullLogPath() As String
    ReDim strFullLogPath(LBound(FileName()) To UBound(FileName()))
    
    ' ------------ Define Log Dirs and such -----------------------------------------
    For A = LBound(FileName()) To UBound(FileName())
        arrLogInfo() = CreateLogFileInfo(FileName(A))
        strStyleDir(A) = arrLogInfo(1)
        strLogDir(A) = arrLogInfo(2)
        strFullLogPath(A) = arrLogInfo(3)
    Next A
    
    'Debug.Print "Style Dir is: " & strStyleDir(1) & vbNewLine & _
                "Log dir is: " & strLogDir(1) & vbNewLine & _
                "Full path to log is: " & strFullLogPath(1)
                
    ' ------------- Check if we need to do an installation ---------------------------
    ' Check if template exists
    Dim installCheck() As Boolean
    ReDim installCheck(LBound(FileName()) To UBound(FileName()))
    Dim blnTemplateExists() As Boolean
    ReDim blnTemplateExists(LBound(FileName()) To UBound(FileName()))
    Dim blnLogUpToDate() As Boolean
    ReDim blnLogUpToDate(LBound(FileName()) To UBound(FileName()))
    Dim logString As String
    Dim strTypeOfInstall As String

    Dim B As Long
    
    For B = LBound(FileName()) To UBound(FileName())
        
        ' Check if log dir/file exists, create if it doesn't, check last mod date if it does
        ' We don't need the true/false info for Installer, but we DO need to run these two
        ' functions to create directories if they don't exist yet
        
        ' If last mod date less than 1 day ago, CheckLog = True
        blnLogUpToDate(B) = CheckLog(strStyleDir(B), strLogDir(B), strFullLogPath(B))
        'Debug.Print FileName(b) & " log exists and was checked today: " & blnLogUpToDate(b)
        
        ' Check if template exists, if not create any missing directories
        blnTemplateExists(B) = IsTemplateThere(FinalDir(B), FileName(B), strFullLogPath(B))
        ' Debug.Print FileName(b) & " exists: " & blnTemplateExists(b)
        
        ' ===============================
        ' FOR DEBUGGING: SET TO TRUE,    |
        ' SO ALWAYS DOWNLOADS FILES      |
        ' Installer = True              '|
        ' ===============================
        
        If Installer = False Then 'Because if it's an installer, we just want to install the file

                
            ' ==========================================
            ' FOR DEBUGGING: SET TO FALSE AND THEN TRUE |
            ' TO TEST NEEDUPDATE FUNCTION               |
            ' blnLogUpToDate(b) = False                '|
            ' blnTemplateExists(b) = True              '|
            ' ==========================================
                
            If blnLogUpToDate(B) = True And blnTemplateExists(B) = True Then ' already checked today, already exists
                installCheck(B) = False
            ElseIf blnLogUpToDate(B) = False And blnTemplateExists(B) = True Then 'Log is new or not checked today, already exists
                'check version number
                installCheck(B) = NeedUpdate(DownloadFrom, FinalDir(B), FileName(B), strFullLogPath(B))
            Else ' blnTemplateExists = False, just download new template
                 installCheck(B) = True
            End If
        Else
            installCheck(B) = True
        End If
        
    Next B

    ' ---------------- Create new array of template files we need to install -----------------
    Dim strInstallFile() As String
    Dim strInstallDir() As String
    Dim C As Long
    Dim X As Long
    
    X = 0
    
    For C = LBound(FileName()) To UBound(FileName())
        If installCheck(C) = True Then
            X = X + 1
            ReDim Preserve strInstallFile(1 To X)
                strInstallFile(X) = FileName(C)
            ReDim Preserve strInstallDir(1 To X)
                strInstallDir(X) = FinalDir(C)
        End If
    Next C
    
    'Debug.Print strInstallFile(1) & vbNewLine & strInstallDir(1)
    
    ' ---------------- Check if new array is allocated -----------------------------------
    If IsArrayEmpty(strInstallFile()) = True Then       ' No files need to be installed
        If Installer = True Then  ' Though this option (no files to install on installer) shouldn't actually occur
            #If Mac Then    ' because application.quit generates error on Mac
                ActiveDocument.Close (wdDoNotSaveChanges)
            #Else
                Application.Quit (wdDoNotSaveChanges)
            #End If
        Else
            Exit Sub
        End If
    Else ' There are values in the array and we need to install them
    
        ' Alert user that installation is happening
        Dim strWelcome As String
        If Installer = True Then
            strWelcome = "Welcome to the " & TemplateName & " Installer!" & vbNewLine & vbNewLine & _
                "Please click OK to begin the installation. It should only take a few seconds."
        Else
            strWelcome = "Your " & TemplateName & " is out of date. Click OK to update automatically."
        End If
    
        If MsgBox(strWelcome, vbOKCancel, TemplateName) = vbCancel Then
            MsgBox "Please try to install the files at a later time."
            
            If Installer = True Then
                ActiveDocument.Close (wdDoNotSaveChanges)
            End If
            
            Exit Sub
        End If
    End If
    
    ' ---------------- Close any open docs (with prompt) -----------------------------------
    Call CloseOpenDocs
        
    '----------------- download template files ------------------------------------------
    Dim D As Long
    
    For D = LBound(strInstallFile()) To UBound(strInstallFile())
    
        If IsReadOnly(strInstallDir(D)) = True Then
            ' Can't replace with new file if destination is read-only; Startup on Mac w/o admin is read-only
            Dim strReadOnlyError As String
            
            strReadOnlyError = "Sorry, you don't have permission to install the file " & strInstallFile(D) & vbNewLine & vbNewLine & _
                "If you are in-house at Macmillan on a Mac, try re-installing the Macmillan Style Template & Macros from the Digital Workflow category in Self Service."
                
                MsgBox strReadOnlyError, vbOKOnly, "Update Failed"
                Exit Sub
        Else
            'If False, error in download; user was notified in DownloadFromConfluence function
            If DownloadFromConfluence(DownloadSource:=DownloadFrom, FinalDir:=strInstallDir(D), _
                LogFile:=strFullLogPath(D), FileName:=strInstallFile(D)) = False Then
                If Installer = True Then
                    #If Mac Then    ' because application.quit generates error on Mac
                        ActiveDocument.Close (wdDoNotSaveChanges)
                    #Else
                        Application.Quit (wdDoNotSaveChanges)
                    #End If
                Else
                    Exit Sub
                End If
            End If
        End If
        
        ' If we just updated the main template, delete the old toolbar
        ' Will be added again by MacmillanGT AutoExec when it's launched, to capture updates
        #If Mac Then
            Dim Bar As CommandBar
            If strInstallFile(D) = "MacmillanGT.dotm" Then
                For Each Bar In CommandBars
                    If Bar.Name = "Macmillan Tools" Then
                        Bar.Delete
                        'Exit For  ' Actually don't exit, in case there are multiple toolbars
                    End If
                    Next
            End If
        #End If
    Next D
    
    '------Display installation complete message   ---------------------------
    Dim strComplete As String
    Dim strInstallType As String
    
    ' Quit if it's an installer, but not if it's an updater (updater was causing conflicts between GT and GtUpdater)
    If Installer = True Then
        strInstallType = "installed"
    Else
        strInstallType = "updated"
    End If
    
    strComplete = "The " & TemplateName & " has been " & strInstallType & " on your computer."
    MsgBox strComplete, vbOKOnly, "Installation Successful"
    
End Sub

Private Function IsTemplateThere(Directory As String, FileName As String, Log As String)
    
    'Create full path to template
    Dim strFullPath As String
    Dim logString As String
    strFullPath = Directory & Application.PathSeparator & FileName
    
    '------------------------- Check if template exists ----------------------------------
    ' Create directory if it doesn't exist
    If IsItThere(Directory) = False Then
        MkDir (Directory)
        IsTemplateThere = False
        logString = Now & " -- Creating template directory."
    Else
        ' Check if template file exists
        If IsItThere(strFullPath) = False Then
            IsTemplateThere = False
            logString = Now & " -- " & FileName & " doesn't exist in " & Directory
        Else
            IsTemplateThere = True
            logString = Now & " -- " & FileName & " already exists."
        End If
    End If

    LogInformation Log, logString
End Function

Private Function NeedUpdate(DownloadURL As GitBranch, Directory As String, FileName As String, Log As String) As Boolean
'Directory argument should be the final directory the template should go in.
'File should be the template file name
'Log argument should be full path to log file

    '------------------------- Get installed version number -----------------------------------
    Dim logString As String
    Dim strFullTemplatePath As String
    
    strFullTemplatePath = Directory & Application.PathSeparator & FileName
    'Debug.Print "NeedUpdate Path: " & strFullTemplatePath
    
    'Get version number of installed template
    Dim strInstalledVersion As String
    
    If IsItThere(strFullTemplatePath) = True Then
        
        #If Mac Then
            Call OpenDocMac(strFullTemplatePath)
        #Else
            Call OpenDocPC(strFullTemplatePath)
        #End If
        
        strInstalledVersion = Documents(strFullTemplatePath).CustomDocumentProperties("Version")
        Documents(strFullTemplatePath).Close SaveChanges:=wdDoNotSaveChanges
        logString = Now & " -- installed version is " & strInstalledVersion
'        Debug.Print "InstalledVersion : |" & strInstalledVersion; "|"
    Else
        strInstalledVersion = "0"     ' Template is not installed
        logString = Now & " -- No template installed, version number is 0."
    End If
    
    LogInformation Log, logString
    
    '------------------------- Try to get current version's number from Confluence ------------
    Dim strVersion As String
    Dim strMacDocs As String
    Dim strStyleDir As String
    Dim strFullVersionPath As String
    
    'Debug.Print FileName
    'Debug.Print InStrRev(FileName, ".do")
    strVersion = Left(FileName, InStrRev(FileName, ".do") - 1)
    strVersion = strVersion & ".txt"
    
    ' Always download version file to Style Directory - on Mac can't write to Startup w/o admin priv
    ' In future, find some way to use this w/o hard-coding the style dir location
    #If Mac Then
        strMacDocs = MacScript("return (path to documents folder) as string")
        strStyleDir = strMacDocs & "MacmillanStyleTemplate"
    #Else
        strStyleDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate"
    #End If
    
    strFullVersionPath = strStyleDir & Application.PathSeparator & strVersion
    'Debug.Print strVersion
    
    'If False, error in download; user was notified in DownloadFromConfluence function
    If DownloadFromConfluence(DownloadSource:=DownloadURL, FinalDir:=strStyleDir, LogFile:=Log, _
        FileName:=strVersion) = False Then
            NeedUpdate = False
            Exit Function
    End If
        
    '-------------------- Get version number of current template ---------------------
    If IsItThere(strFullVersionPath) = True Then
        NeedUpdate = True
        Dim strCurrentVersion As String

        strCurrentVersion = ReadTextFile(Path:=strFullVersionPath, FirstLineOnly:=True)
        
        ' git converts all line endings to LF which messes up PC, and I don't want to deal
        ' with it so we'll just remove everything
        If InStr(strCurrentVersion, vbCrLf) > 0 Then
            strCurrentVersion = Replace(strCurrentVersion, vbCrLf, "")
        ElseIf InStr(strCurrentVersion, vbLf) > 0 Then
            strCurrentVersion = Replace(strCurrentVersion, vbLf, "")
        ElseIf InStr(strCurrentVersion, vbCr) > 0 Then
            strCurrentVersion = Replace(strCurrentVersion, vbCr, "")
        End If
        
'        Debug.Print "Text File: |" & strCurrentVersion & "|"

        logString = Now & " -- Current version is " & strCurrentVersion
    Else
        NeedUpdate = False
        logString = Now & " -- Download of version file for " & FileName & " failed."
    End If
        
    LogInformation Log, logString
    
    If NeedUpdate = False Then
        Exit Function
    End If
    
    '--------------------- Compare version numbers -----------------------------------
    
    If strInstalledVersion >= strCurrentVersion Then
        NeedUpdate = False
        logString = Now & " -- Current version matches installed version."
    Else
        NeedUpdate = True
        logString = Now & " -- Current version greater than installed version."
    End If
    
    LogInformation Log, logString
    
End Function

Private Sub OpenDocMac(FilePath As String)
        Documents.Open FileName:=FilePath, ReadOnly:=True ', Visible:=False      'Mac Word 2011 doesn't allow Visible as an argument :(
End Sub

Private Sub OpenDocPC(FilePath As String)
        Documents.Open FileName:=FilePath, ReadOnly:=True, Visible:=False      'Win Word DOES allow Visible as an argument :)
End Sub



Private Function ImportVariable(strFile As String) As String
 
    Open strFile For Input As #1
    Line Input #1, ImportVariable
    Close #1
 
End Function



