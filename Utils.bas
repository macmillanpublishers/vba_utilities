Attribute VB_Name = "GeneralHelpers"

' All should be declared as Public for use from other modules

' *****************************************************************************
'           DECLARATIONS
' *****************************************************************************
Option Explicit
Private Const strModule As String = "genUtils.GeneralHelpers."
Public lngErrorCount As Long


' assign to actual document we're working on
' to do: probably better managed via a class
Public activeDoc As Document

' *****************************************************************************
'           ERROR HANDLING STUFF
' *****************************************************************************
' Technically class methods (which are listed farther down) but it's easier to
' maintain close to the error number enum, which has to appear above other
' procedures. May move to its own class one day.

' ===== MacErrors =========================================================
' Each error we raise anywhere in this class should be included here. Note that
' error numbers 0 - 512 are reserved for system errors. Include property or
' method name in enum just to be clear for debugging.

Public Enum MacError
    err_MacErrGeneral = 20000       ' Not sure if we'll need, but here it is.
    err_GroupNameInvalid = 20001
    err_GroupNameNotSet = 20002
    err_SpecificFileInvalid = 20003
    err_SpecificFileNotSet = 20004
    err_DeleteThisDoc = 20005
    err_TempDeleteFail = 20006
    err_NoInternet = 20007
    err_Http404 = 20008
    err_BadHttpStatus = 20009
    err_DownloadFail = 20010
    err_LocalDeleteFail = 20011
    err_LocalCopyFail = 20012
    err_LocalReadOnly = 20013
    err_TempReadOnly = 20014
    err_TempMissing = 20015
    err_FileNotThere = 20016
    err_NotWordFormat = 20017
    err_ConfigPathNull = 20018
    err_RootDirInvalid = 20019
    err_RootDirMissing = 20020
    err_LogReadOnly = 20021
    err_DirectoryMissing = 20022
    err_ParaIndexInvalid = 20023
    err_BacktickCharFound = 20024
    err_DocProtectionOn = 20025
    err_NotArray = 20026
End Enum

' ===== ErrorChecker ==========================================================
' Send all errors here, to make tracking/maintaining error handling easier.
' Be sure to add the error to the MacErrors enum.

' DO NOT use any other of our functions in this function, because those need to
' direct errors here and using them here as well could create an infinite loop.
'
' USE: ErrorChecker() returns False if error is handled; procedure can continue

Public Function ErrorChecker(objError As Object, Optional strValue As _
    String) As Boolean
    ' strValue - varies based on type of error passed. use for things like
    ' file name, path, whatever is being checked by that errored.
  lngErrorCount = lngErrorCount + 1
  DebugPrint "ErrorChecker " & lngErrorCount & vbNewLine & _
    "(" & objError.Source & ") " & objError.Number & ":" & vbNewLine _
    & objError.Description
  
  If lngErrorCount > 15 Then
    DebugPrint "ERROR LOOP STOPPED"
    End
  End If
  
  ' New On Error statement RESETS the Err object, so get our values out before
  ' we set the ErrorChecker for this procedure.
  Dim lngErrNumber As Long
  Dim strErrDescription As String
  Dim strErrSource As String
  
  lngErrNumber = objError.Number
  strErrDescription = objError.Description  ' For system errors
  strErrSource = objError.Source
  
'  On Error GoTo ErrorCheckerError

  Dim strErrMessage As String
  Dim blnNotifyUser As Boolean
  Dim strHelpContact As String

  ' ----- Set defaults --------------------------------------------------------
  ' Only need to change in Select statement below if want to set either to
  ' False.
  blnNotifyUser = True
  ErrorChecker = True
  ' Eventually get this email from the config.json file, once we have error
  ' handling that can handle an error in the error handler. At the moment, if
  ' any procedure in the call stack to get the email errors, we end up in a
  ' loop until it crashes.
  strHelpContact = vbNewLine & vbNewLine & "Email workflows@macmillan.com if" _
    & " you need help. Be sure to attach the text file that was just produced."

  ' ----- Check if parameter was passed ---------------------------------------
  If strValue = vbNullString Then
    strValue = "UNKNOWN"
  End If

  ' ----- Check for errors ----------------------------------------------------
  ' Make sure we actually have an error, cuz you never know.
  If lngErrNumber <> 0 Then

  ' ----- Handle specific errors ----------------------------------------------
  ' Each case should be a MacError enum. Even if we aren't doing anything
  ' to fix the error, still enter the following:
  '   strErrDescription = "Description of the error for the log."
  '   strErrMessage = "Message for the user if we're notifying them."
  Select Case lngErrNumber
  
    Case 5941, 5834
    ' 5941: Item not present in collection
    ' 5834: Item with the specified name does not exist
    ' Most common cause of these is referencing a style that doesn't exist.
    
    ' BUT: How can we be sure this means *style* is not present?
    
      'Anyway, If style is not present, add style
      Dim myStyle As Style
      Dim styleType As WdStyleType
      If InStr(strValue, "span") > 0 Then
        styleType = wdStyleTypeCharacter
      Else
        styleType = wdStyleTypeParagraphOnly
      End If
      Set myStyle = activeDoc.Styles.Add(strValue, styleType)
      ErrorChecker = False
      DebugPrint "ErrorChecker: False"
      Exit Function
    ' List all built-in errors we want to trap for before general sys error line
    Case 91 ' Object variable or With block variable not set.
      ' May be caused if `activeDoc` global var is not set
      If activeDoc Is Nothing Then
        Set activeDoc = ActiveDocument
        ErrorChecker = False
        DebugPrint "ErrorChecker: False"
        Exit Function
      End If
      
    Case Is < 513
      ' Errors 0 to 512 are system errors
      strErrMessage = "Something unexpected happened. Please click OK to exit." _
        & vbNewLine & "Value: " & strValue
    Case MacError.err_GroupNameInvalid
      strErrDescription = "Invalid value for GroupName property: " & strValue
      strErrMessage = "The value you've entered for the GroupName property, " _
        & "is not valid. Make sure you only use file groups that are in the " _
        & "config.json file."
    Case MacError.err_GroupNameNotSet
      strErrDescription = "GroupName property has not been Let."
      strErrMessage = "You can't Get the GroupName property before it has " _
        & "been Let. Try the MacFile_.AssignFile method to create a new object in this class."
    Case MacError.err_SpecificFileInvalid
      strErrDescription = "Invalid value for SpecificFile property: " _
          & strValue
      strErrMessage = "The you've entered an invalid value for the" _
          & " SpecificFile property. Make sure you only use " & _
          "specific file types that are in the config.json file."
    Case MacError.err_SpecificFileNotSet
      strErrDescription = "SpecificFile property has not been Let."
      strErrMessage = "You can't Get the SpecificFile property " & _
          "before it has been Let. Try the MacFile_.AssignFile " & _
          "method to create a new object in this class."
    Case MacError.err_DeleteThisDoc
      strErrDescription = "Can't delete file that is currently " & _
          "executing code: " & strValue
      strErrMessage = "The file you are trying to delete is " & _
          "currently executing macro code."
    Case MacError.err_TempDeleteFail
      strErrDescription = "Failed to delete the previous file in the " _
          & "temp directory: " & strValue
      strErrMessage = "We can't download the file; a temp file " & _
          "is still there."
    Case MacError.err_NoInternet
      strErrDescription = "No network connection. Download aborted."
      strErrMessage = "We weren't able to download the file " & _
          "because we can't connect to the internet. Check your " & _
          "network connection and try again."
    Case MacError.err_Http404
      strErrDescription = "File HTTP status 404. Check if DownloadURL" _
          & " is correct, and file is posted: " & strValue
      strErrMessage = "Could not download file from the internet."
    Case MacError.err_BadHttpStatus
      strErrDescription = "File HTTP status: " & strValue & _
            ". Download aborted."
      strErrMessage = "There is some problem with the file you are" _
            & " trying to download."
        ' Need to get Source as passed in object first, so do this last
    Case MacError.err_DownloadFail
      strErrDescription = "File download failed: " & strValue
      strErrMessage = "Download failed."
    Case MacError.err_LocalDeleteFail
        ' genUtils.GeneralHelpers.KillAll() will notify user if file is open
      strErrDescription = "File in final install location could not be " & _
        "deleted. If the file was open, the user was notified: " & strValue
      blnNotifyUser = False
    Case MacError.err_LocalCopyFail
      strErrDescription = "File not saved to final directory: " & strValue
      strErrMessage = "There was an error installing the Macmillan template."
    Case MacError.err_LocalReadOnly
      strErrDescription = "Final dir for file is read-only: " & strValue
      strErrMessage = "The folder you are trying to access is read-only."
    Case MacError.err_TempReadOnly
      strErrDescription = "Temp dir is read-only: " & strValue
      strErrMessage = "Your temp folder is read-only."
    Case MacError.err_TempMissing
      strErrDescription = "Temp directory is missing: " & strValue
       strErrMessage = "There is an error with your temp folder."
    Case MacError.err_FileNotThere
       strErrDescription = "File does not exist: " & strValue
       strErrMessage = "The file " & strValue & " does " _
            & "not exist."
    Case MacError.err_NotWordFormat
      strErrDescription = "File extension is not a native Word " & _
            "document or template: " & strValue
       strErrMessage = "This file does not appear to be a Word " & _
            "file: " & strValue
    Case MacError.err_ConfigPathNull
       strErrDescription = "FullConfigPath custom doc property is not " _
            & "set in the document."
      strErrMessage = "We can't find the config.json file because " _
            & "the local path is not in the template properties."
    Case MacError.err_RootDirInvalid
      strErrDescription = "Value for root directory in config.json is" _
            & " not an option in the RootDir property: " & strValue
      strErrMessage = "The folder where we save the Tools template" _
          & " doesn't exist."
    Case MacError.err_LogReadOnly
      strErrDescription = "Log file is read only: " & strValue
      strErrMessage = "There is a problem with the logs."
    Case MacError.err_DirectoryMissing
      strErrDescription = "The directory " & strValue & " is missing."
      strErrMessage = strErrDescription
    Case MacError.err_ParaIndexInvalid
      strErrDescription = "The requested paragraph is out of range."
      strErrMessage = strErrDescription
    Case MacError.err_BacktickCharFound
      strErrDescription = "Backtick (`) character found in manuscript. A " & _
        "macro was probably run before and failed."
      strErrMessage = strErrDescription
    Case MacError.err_DocProtectionOn
      strErrDescription = "Document protection is enabled. Ask original user" _
        & " to unlock the file and try again."
      strErrMessage = strErrDescription
    Case MacError.err_NotArray
      strErrDescription = "Variable is not an array."
      strErrMessage = strErrDescription
    Case Else
      strErrDescription = "Undocumented error - " & strErrDescription
      strErrMessage = "Not sure what's going on here."
  End Select

  Else
      strErrDescription = "Everything's A-OK. Why are you even reading this?"
      blnNotifyUser = False
      ErrorChecker = False
  End If

  ' ----- WRITE ERROR LOG ---------------------------------------------------
  ' Output text file with error info, user could send via email.
  ' Do not use WriteToLog function, because that sends errors here as well.

  Dim strErrMsg As String
  Dim LogFileNum As Long
  Dim strTimeStamp As String
  Dim strErrLog As String
  Dim strFileName As String

' Check activeDoc:
  If activeDoc Is Nothing Then
    Set activeDoc = ActiveDocument
  End If

  ' Write error log to same location as current file.
  ' Format date so it can be part of file name. Only including date b/c users
  ' will likely run things repeatedly before asking for help, and don't want
  ' to generate a bunch of files if include time as well.
  strFileName = Replace(Right(ActiveDocument.Name, InStrRev(activeDoc.Name, _
    ".") - 1), " ", "")
  strErrLog = ActiveDocument.Path & Application.PathSeparator & "ALERT_" & _
    strFileName & "_" & Format(Date, "yyyy-mm-dd") & ".txt"
'    DebugPrint strErrLog
  ' build error message, including timestamp
  strErrMsg = Format(Time, "hh:mm:ss - ") & strErrSource & vbNewLine & _
      lngErrNumber & ": " & strErrDescription & vbNewLine
  LogFileNum = FreeFile ' next file number
  Open strErrLog For Append As #LogFileNum ' creates the file if doesn't exist
  Print #LogFileNum, strErrMsg ' write information to end of the text file
  Close #LogFileNum ' close the file
  
  ' Do not display alerts for Bookmaker project (automated)
  If Left(Err.Source, InStr(Err.Source, ".") - 1) = "Bookmaker" Then
    blnNotifyUser = False
  End If
  
'  If blnNotifyUser = True Then
'      strErrMessage = strErrMessage & vbNewLine & vbNewLine & strHelpContact
'      MsgBox Prompt:=strErrMessage, Buttons:=vbExclamation, Title:= _
'          "Macmillan Tools Error"
'  End If
  DebugPrint "ErrorChecker: " & ErrorChecker
  Exit Function

ErrorCheckerError:
  ' Important note: Recursive error checking is perhaps a bad idea -- if the
  ' same error gets triggered, procedure will get called too many times and
  ' cause an "out of stack space" error and crash.
  DebugPrint Err.Number & ": " & Err.Description
  ErrorChecker = True
End Function

' ===== GlobalCleanup =========================================================
' A variety of resetting/cleanup functions

Sub GlobalCleanup()
  On Error GoTo GlobalCleanupError
  GeneralHelpers.zz_clearFind
  If Not activeDoc Is Nothing Then
    Set activeDoc = Nothing
  End If
  Application.DisplayAlerts = wdAlertsAll
  Application.ScreenUpdating = True
  Application.ScreenRefresh
  On Error GoTo 0

GlobalCleanupError:
  ' Halts ALL execution, resets all variables, unloads all userforms, etc.
  End
End Sub



' ===== DebugPrint =============================================================
' Use instead of `Debug.Print`. Print to Immediate Window AND write to a file.
' Immediate Window has a small buffer and isn't very useful if you are debugging
' something that ends up crashing the app.

' Actual `Debug.Print` can take more complex arguments but here we'll just take
' anything that can evaluate to a string.

' Need to set "VbaDebug" environment variable to True also

Public Sub DebugPrint(Optional StringExpression As Variant)

  If Environ("VbaDebug") = True Then
  ' First just DebugPrint:
  ' Get the string we'll write
    Dim strMessage As String
    strMessage = Now & ": " & StringExpression
    Debug.Print strMessage
  
  ' Second, write to file
  ' Create file name
  ' !!! ActiveDocument.Path sometimes writes to STARTUP dir. Also if running
  ' with Folder Actions (like Validator), new file in dir will error
  ' How to write to a static location?
    Dim strOutputFile As String
    strOutputFile = Environ("USERPROFILE") & Application.PathSeparator & _
      "Desktop" & Application.PathSeparator & "immediate_window.txt"
  
    Dim FileNum As Integer
    FileNum = FreeFile ' next file number
    Open strOutputFile For Append As #FileNum
    Print #FileNum, strMessage
    Close #FileNum ' close the file
  End If
 
End Sub


Public Function IsOldMac() As Boolean
  On Error GoTo IsOldMacError
    ' Checks this is a Mac running Office 2011 or earlier. Good for things like
    ' checking if we need to account for file paths > 3 char (which 2011 can't
    ' handle but Mac 2016 can.
    IsOldMac = False
    #If Mac Then
        If Application.Version < 16 Then
            IsOldMac = True
        End If
    #End If
  Exit Function
  
IsOldMacError:
  Err.Source = strModule & "IsOldMac"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function

Public Function DocPropExists(objDoc As Document, PropName As String) As Boolean
  On Error GoTo DocPropExistsError
    ' Tests if a particular custom document property exists in the document. If
    ' it's already a Document object we already know that it exists and is open
    ' so we don't need to test for those here. Should be tested somewhere in
    ' calling procedure though.
    DocPropExists = False

    Dim A As Long
    Dim docProps As DocumentProperties
    docProps = objDoc.CustomDocumentProperties

    If docProps.Count > 0 Then
        For A = 1 To docProps.Count
            If dopProps.Name = PropName Then
                DocPropExists = True
                Exit Function
            End If
        Next A
    Else
        DocPropExists = False
    End If
    Exit Function
DocPropExistsError:
  Err.Source = strModule & "DocPropExists"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function

Public Function IsOpen(DocPath As String) As Boolean
    ' Tests if the Word document is currently open.
    On Error GoTo IsOpenError
    Dim objDoc As Document
    IsOpen = False
    If IsItThere(DocPath) = True Then
        If IsWordFormat(DocPath) = True Then
            If Documents.Count > 0 Then
                For Each objDoc In Documents
                    If objDoc.FullName = DocPath Then
                        IsOpen = True
                        Exit Function
                    End If
                Next objDoc
            End If
        Else
            Err.Raise MacError.err_NotWordFormat
        End If
    Else
        Err.Raise MacError.err_FileNotThere
    End If
    Exit Function

IsOpenError:
    Err.Source = Err.Source & strModule & "IsOpen"
    If ErrorChecker(Err, DocPath) = False Then
        Resume
    Else
        IsOpen = False
        Call genUtils.GeneralHelpers.GlobalCleanup
    End If
End Function

Public Function IsWordFormat(PathToFile As String) As Boolean
  On Error GoTo IsWordFormatError
    ' Checks extension to see if file is a Word document or template. Notably,
    ' does not test if it's a file type that Word CAN open (e.g., .html), just
    ' if it's a native Word file type.

    ' Ignores final character for newer file types, just checks for .dot / .doc
    Dim strExt As String
    strExt = Left(Right(PathToFile, InStr(StrReverse(PathToFile), ".")), 4)
    If strExt = ".dot" Or strExt = ".doc" Then
        IsWordFormat = True
    Else
        IsWordFormat = False
    End If
    Exit Function
IsWordFormatError:
    Err.Source = Err.Source & strModule & "IsWordFormat"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        Call genUtils.GeneralHelpers.GlobalCleanup
    End If
    
End Function

Public Function IsLocked(FilePath As String) As Boolean
    ' Tests if any file is locked by some kind of process.
    On Error GoTo IsLockedError
    IsLocked = False
    If IsItThere(FilePath) = False Then
        Err.Raise MacError.err_FileNotThere
    Else
        Dim FileNum As Long
        FileNum = FreeFile()
        ' If the file is already in use, next line will raise an error:
        ' "70: Permission denied" (file is open, Word doc is loaded as add-in)
        ' "75: Path/File access error" (File is read-only, etc.)
        Open FilePath For Binary Access Read Write Lock Read Write As FileNum
        Close FileNum
    End If
IsLockedFinish:
    Exit Function
    
IsLockedError:
    Err.Source = Err.Source & strModule & "IsLocked"
    If Err.Number = 70 Or Err.Number = 75 Then
        IsLocked = True
        Resume IsLockedFinish
    Else
        If ErrorChecker(Err, FilePath) = False Then
            Resume
        Else
            Call genUtils.GeneralHelpers.GlobalCleanup
        End If
    End If
End Function

Public Function IsItThere(Path As String) As Boolean
  On Error GoTo IsItThereError
    ' Check if file or directory exists on PC or Mac.
    ' Dir() doesn't work on Mac 2011 if file is longer than 32 char
    'DebugPrint Path
    
    'Remove trailing path separator from dir if it's there
    If Right(Path, 1) = Application.PathSeparator Then
        Path = Left(Path, Len(Path) - 1)
    End If

    If IsOldMac = True Then
        Dim strScript As String
        strScript = "tell application " & Chr(34) & "System Events" & Chr(34) & _
            "to return exists disk item (" & Chr(34) & Path & Chr(34) _
            & " as string)"
        IsItThere = genUtils.GeneralHelpers.ShellAndWaitMac(strScript)
    Else
        Dim strCheckDir As String
        strCheckDir = Dir(Path, vbDirectory)
        
        If strCheckDir = vbNullString Then
            IsItThere = False
        Else
            IsItThere = True
        End If
    End If
    Exit Function
IsItThereError:
    Err.Source = Err.Source & strModule & "IsItThere"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        Call genUtils.GeneralHelpers.GlobalCleanup
    End If
End Function

' ===== ParentDirExists =======================================================
' If `FilePath` is the full path to a file (that may or may not exist), then
' this checks that the directory the file is in exists. Good for checking paths
' to files before you create them.

Public Function ParentDirExists(FilePath As String) As Boolean
  On Error GoTo ParentDirExistsError
  Dim strDir As String
  Dim strFile As String
  Dim lngSep As Long
  ParentDirExists = False
  ' Separate directory from file name
  lngSep = InStrRev(FilePath, Application.PathSeparator)
  
  If lngSep > 0 Then
    strDir = VBA.Left(FilePath, lngSep - 1)  ' NO trailing separator
    strFile = VBA.Right(FilePath, Len(FilePath) - lngSep)
'    DebugPrint strDir & " | " & strFile

    ' Verify file name string is in fact plausibly a file name
    If InStr(strFile, ".") > 0 Then
      ' NOW we can check if the directory exists:
      ParentDirExists = IsItThere(strDir)
      Exit Function
    End If
  End If
  Exit Function
ParentDirExistsError:
  Err.Source = Err.Source & strModule & "ParentDirExists"
  If ErrorChecker(Err, FilePath) = False Then
    Resume
  Else
     Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function

Public Function KillAll(Path As String) As Boolean
    ' Deletes file (or folder?) on PC or Mac. Mac can't use Kill() if file name
    ' is longer than 32 char. Returns true if successful.
    On Error GoTo KillAllError
    If IsItThere(Path) = True Then
        ' Can't delete file if it's installed as an add-in
        If IsInstalledAddIn(Path) = True Then
            AddIns(Path).Installed = False
        End If
        ' Mac 2011 can't handle file paths > 32 char
        #If Mac Then
            If Application.Version < 16 Then
                Dim strCommand As String
                strCommand = MacScript("return quoted form of posix path of " & Path)
                strCommand = "rm " & strCommand
                genUtils.GeneralHelpers.ShellAndWaitMac (strCommand)
            Else
                Kill (Path)
            End If
        #Else
            Kill (Path)
        #End If

        ' Make sure it worked
        If IsItThere(Path) = False Then
            KillAll = True
        Else
            KillAll = False
        End If
    Else
        KillAll = True
    End If
KillAllFinish:
    Exit Function
    
KillAllError:
    Dim strErrMsg As String
    Select Case Err.Number
        Case 70     ' File is open
            strErrMsg = "Please close all other Word documents and try again."
            MsgBox strErrMsg, vbCritical, "Macmillan Tools Error"
            KillAll = False
            Resume KillAllFinish
        Case Else
          Err.Source = strModule & "KillAll"
          If genUtils.GeneralHelpers.ErrorChecker(Err, strLogFile) = False Then
            Resume
          Else
            Call genUtils.GeneralHelpers.GlobalCleanup
          End If
    End Select
End Function

Public Function IsInstalledAddIn(FileName As String) As Boolean
  On Error GoTo IsInstalledAddInError
    ' Check if the file is currently loaded as an AddIn. Because we can't delete
    ' it if it is loaded (though we can delete it if it's just referenced but
    ' not loaded).
    Dim objAddIn As AddIn
    For Each objAddIn In AddIns
        ' Check if in collection first; throws error if try to check .Installed
        ' but it's not even referenced.
        If objAddIn.Name = FileName Then
            If objAddIn.Installed = True Then
                IsInstalledAddIn = True
            Else
                IsInstalledAddIn = False
            End If
            Exit For
        End If
    Next objAddIn
  Exit Function
IsInstalledAddInError:
  Err.Source = Err.Source & strModule & "IsInstalledAddIn"
  If ErrorChecker(Err) = False Then
    Resume
  Else
     Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function

' ===== WriteToLog ============================================================
' Writes line to log for the file. LogMessage only needs text, timestamp will
' be added in this method.

Public Sub WriteToLog(LogMessage As String, Optional LogFilePath As String)
    On Error GoTo WriteToLogError
    Dim strLogFile As String
    Dim strLogMessage As String
    Dim FileNum As Integer

    ' If no specific path was passed, write to generic log file
    If LogFilePath = vbNullString Then
        strLogFile = Paths_.LogsDir & Application.PathSeparator & _
            "manuscript-tools.log"
    Else
        strLogFile = LogFilePath
    End If

    If IsItThere(strLogFile) = True Then
        If IsReadOnly(strLogFile) = True Then
            Err.Raise MacError.err_LogReadOnly
        End If
    End If

    ' prepend current date and time to message
    strLogMessage = Now & " -- " & LogMessage
    FileNum = FreeFile ' next file number
    Open strLogFile For Append As #FileNum ' creates the file if doesn't exist
    Print #FileNum, strLogMessage ' write information to end of the text file
    Close #FileNum ' close the file
WriteToLogFinish:
    Exit Sub

WriteToLogError:
    Err.Source = Err.Source & strModule & "WriteToLog"
    If genUtils.GeneralHelpers.ErrorChecker(Err, strLogFile) = False Then
        Resume
    Else
        Call genUtils.GeneralHelpers.GlobalCleanup
    End If
End Sub

Public Function ShellAndWaitMac(cmd As String) As String
    Dim result As String
    Dim scriptCmd As String ' Macscript command
    #If Mac Then
        scriptCmd = "do shell script " & Chr(34) & cmd & Chr(34) & Chr(34)
        result = MacScript(scriptCmd) ' result contains stdout, should you care
        'DebugPrint result
        ShellAndWaitMac = result
    #End If
End Function

Public Function IsStyleInUse(StyleName As String) As Boolean
  On Error GoTo IsStyleInUseError
  
' First confirm style is even in document to begin with
  If GeneralHelpers.IsStyleInDoc(StyleName) = False Then
    IsStyleInUse = False
    Exit Function
  End If

'  ' If we need to do a Selection.Find use
'  Selection.HomeKey Unit:=wdStory
  Call genUtils.zz_clearFind
  With activeDoc.Range.Find
    .Text = ""
    .Format = True
    .Style = activeDoc.Styles(StyleName)
    .Execute
    
    If .Found = True Then
      IsStyleInUse = True
    Else
      IsStyleInUse = False
    End If
  End With
  
  Exit Function
IsStyleInUseError:
  Err.Source = strModule & "IsStyleInUse"
  If ErrorChecker(Err, StyleName) = False Then
    Resume
  Else
    Call genUtils.GlobalCleanup
  End If
End Function


Public Function IsStyleInDoc(StyleName As String) As Boolean
  On Error GoTo IsStyleInDocError
  Dim blnResult As Boolean: blnResult = True
  Dim TestStyle As Style
  
' Try to access this style. If not present in doc, will error
  Set TestStyle = activeDoc.Styles.Item(StyleName)
  IsStyleInDoc = blnResult
  Exit Function
  
IsStyleInDocError:
' 5941 = "The requested member of the collection does not exist."
' Have to test here, ErrorChecker tries to create style if missing
  If Err.Number = 5941 Then
    blnResult = False
    Resume Next
  End If
' Otherwise, usual error stuff:
  Err.Source = strModule & "IsStyleInDoc"
  If ErrorChecker(Err, StyleName) = False Then
    Resume
  Else
    Call genUtils.GlobalCleanup
  End If
End Function

Public Sub OverwriteTextFile(TextFile As String, NewText As String)
' TextFile should be full path
  On Error GoTo OverwriteTextFileError
  Dim FileNum As Integer
  ' Will create file if not exist, but parent dir must exist
  If ParentDirExists(TextFile) = True Then
    FileNum = FreeFile ' next file number
    Open TextFile For Output Access Write As #FileNum
    Print #FileNum, NewText ' overwrite information in the text of the file
    Close #FileNum ' close the file
  Else
    ' directory is invalid
    Err.Raise MacError.err_DirectoryMissing
  End If
  Exit Sub
  
OverwriteTextFileError:
  Err.Source = strModule & "OverwriteTextFile"
  If ErrorChecker(Err, TextFile) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub


Public Sub AppendTextFile(TextFile As String, Contents As String)
' TextFile should be full path
  On Error GoTo AppendTextFileError
  Dim FileNum As Integer
' Will create file if not exist, but parent dir must exist
  TextFile = VBA.Replace(TextFile, "/", Application.PathSeparator)
  If ParentDirExists(TextFile) = True Then
    FileNum = FreeFile ' next file number
    Open TextFile For Append As #FileNum
    Print #FileNum, Contents
    Close #FileNum ' close the file
  Else
    ' directory is invalid
    Err.Raise MacError.err_DirectoryMissing
  End If
  Exit Sub
  
AppendTextFileError:
  Err.Source = strModule & "AppendTextFile"
  If ErrorChecker(Err, TextFile) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Public Function SetPathSeparator(strOrigPath As String) As String
' Must pass full path, throws error if no path separators found.
  On Error GoTo SetPathSeparatorError
  Dim strFinalPath As String
  strFinalPath = strOrigPath
  
  Dim strCharacter(1 To 3) As String
  strCharacter(1) = ":"
  strCharacter(2) = "/"
  strCharacter(3) = "\"
  
  Dim A As Long
  For A = LBound(strCharacter) To UBound(strCharacter)
    If InStr(strOrigPath, A) > 0 Then
      strFinalPath = VBA.Replace(strFinalPath, A, Application.PathSeparator)
    End If
  Next A
  
  SetPathSeparator = strFinalPath
  Exit Function
  
SetPathSeparatorError:
  Err.Source = strModule & "SetPathSeparator"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function

Public Function CheckLog(StyleDir As String, LogDir As String, LogPath As String) As Boolean
'LogPath is *full* path to log file, including file name. Created by CreateLogFileInfo sub, to be called before this one.

    Dim logString As String
    
    '------------------ Check log file --------------------------------------------
    'Check if logfile/directory exists
    If IsItThere(LogPath) = False Then
        CheckLog = False
        logString = Now & " -- Creating logfile."
        If IsItThere(LogDir) = False Then
            If IsItThere(StyleDir) = False Then
                MkDir (StyleDir)
                MkDir (LogDir)
                logString = Now & " -- Creating MacmillanStyleTemplate directory."
            Else
                MkDir (LogDir)
                logString = Now & " -- Creating log directory."
            End If
        End If
    Else    'logfile exists, so check last modified date
        Dim lastModDate As Date
        lastModDate = FileDateTime(LogPath)
        If DateDiff("d", lastModDate, Date) < 1 Then       'i.e. 1 day
            CheckLog = True
            logString = Now & " -- Already checked less than 1 day ago."
        Else
            CheckLog = False
            logString = Now & " -- >= 1 day since last update check."
        End If
    End If
    
    'Log that info!
    LogInformation LogPath, logString
    
End Function

' ===== ParaIndex =============================================================
' Returns the paragraph index of the current selection. Default is to return the
' END paragraph index if selection is more than 1 paragraph. `UseEnd:=False`
' would return the index of the START paragraph.
Public Function ParaIndex(Optional UseEnd As Boolean = True) As Long
  On Error GoTo ParaIndexError
  If UseEnd = True Then
    ParaIndex = activeDoc.Range(0, Selection.End).Paragraphs.Count
  Else
    ParaIndex = activeDoc.Range(0, Selection.Start).Paragraphs.Count
  End If
  Exit Function
ParaIndexError:
  Err.Source = strModule & "ParaIndex"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    genUtils.GlobalCleanup
  End If
End Function

' ===== ParaInfo ==============================================================
' In general: get a variety of info about a paragraph (or its document)
' Most common usage: InfoType = wdActiveEndAdjustedPageNumber

Public Function ParaInfo(paraInd As Long, InfoType As WdInformation) _
  As Variant
  On Error GoTo ParaInfoError
  
' Make sure we have an activeDoc
  If activeDoc Is Nothing Then
    Set activeDoc = ActiveDocument
  End If
  
' Make sure our paragraph index is in range
  If paraInd <= activeDoc.Paragraphs.Count Then
  ' Set range for our paragraph
    Dim rngPara As Range
    Set rngPara = activeDoc.Paragraphs(paraInd).Range
    ParaInfo = rngPara.Information(InfoType)
  Else
    Err.Raise MacError.err_ParaIndexInvalid
  End If

  Exit Function
ParaInfoError:
  Err.Source = strModule & "ParaInfo"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call GlobalCleanup
  End If
End Function

Public Sub zz_clearFind()
  On Error GoTo zz_clearFindError
'  lngErrorCount = lngErrorCount + 1
'  DebugPrint "zz_clearFind " & lngErrorCount
    Dim clearRng As Range
    Set clearRng = activeDoc.Words.First
'    DebugPrint activeDoc.FullName
    
    With clearRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " "
        .Replacement.Text = " "
        .Wrap = wdFindStop
        .Format = False
        .Forward = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceOne
    End With
  Exit Sub
zz_clearFindError:
' Can't do any replace if doc is password protected, but this runs
' as part of cleanup so need to handle that here:
  If Err.Number = 9099 Then ' "Command is not available"
    Exit Sub
  End If
  Err.Source = strModule & "zz_clearFind"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call GlobalCleanup
  End If
End Sub

Public Function StoryArray() As Variant
  On Error GoTo StoryArrayError
    '------------check for endnotes and footnotes--------------------------
    Dim strStories() As Variant
    
    ReDim strStories(1 To 1)
    strStories(1) = wdMainTextStory
    
    If ActiveDocument.Endnotes.Count > 0 Then
        ReDim Preserve strStories(1 To (UBound(strStories()) + 1))
        strStories(UBound(strStories())) = wdEndnotesStory
    End If
    
    If ActiveDocument.Footnotes.Count > 0 Then
        ReDim Preserve strStories(1 To (UBound(strStories()) + 1))
        strStories(UBound(strStories())) = wdFootnotesStory
    End If
    
    StoryArray = strStories
  Exit Function
StoryArrayError:
  Err.Source = strModule & "StoryArray"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call GlobalCleanup
  End If
End Function

Function PatternMatch(SearchPattern As String, SearchText As String, WholeString As Boolean) As Boolean
    ' "SearchPattern" uses Word Find pattern matching, which is not the same as regular expressions
    ' But the RegEx library breaks Word Mac 2011, so we'll do it this way
    ' This is a good reference: http://www.gmayor.com/replace_using_wildcards.htm
    ' "SearchText" is the string you're looking in
    ' "WholeString" is True if you are trying to match the whole string; if just part
    ' of the string is an acceptable match, set to False
        
    ' Need to paste string into a Word doc to use Find pattern matching
    Dim newDoc As New Document
    Set newDoc = Documents.Add(Visible:=False)
    newDoc.Select
    
    Selection.InsertBefore (SearchText)
    ' Insertion point has to be at start of doc for Selection.Find
    Selection.Collapse (wdCollapseStart)
    
    With Selection.Find
        .ClearFormatting
        .Text = SearchPattern
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWholeWord = False
        .MatchCase = True
        .MatchWildcards = True
        .MatchSoundsLike = False
        .Execute
    End With
    
    If Selection.Find.Found = True Then
        If WholeString = True Then
            ' The final paragraph return is the only character the new doc had it in,
            ' it 's not part of the added string
            If InStrRev(Selection.Text, Chr(13)) = Len(Selection.Text) Then
                Selection.MoveEnd Unit:=wdCharacter, Count:=-1
            End If
            
            ' the SearchText requires vbCrLf to start text on a new line, but Word for some reason
            ' strips out the Lf when content is pasted in. CrLf counts as 2 characters but Cr is only
            ' 1, so to get these to match we need to add 1 character to the selection for each line.
            Dim lngLines As Long
            lngLines = ActiveDocument.ComputeStatistics(wdStatisticLines)
            
            If Len(Selection.Text) + lngLines = Len(SearchText) Then
                PatternMatch = True
            Else
                PatternMatch = False
            End If
        Else
            PatternMatch = True
        End If
    Else
        PatternMatch = False
    End If
    
    newDoc.Close wdDoNotSaveChanges
    
End Function

Function CheckSave()
'    ' Prompts user to save document before running the macro. If they click "Cancel" then CheckSave returns true and
'    ' you should exit your macro. also checks if document protection is on.
'
'    Dim mainDoc As Document
'    Set mainDoc = ActiveDocument
'    Dim iReply As Integer
'
'    '-----make sure document is saved
'    Dim docSaved As Boolean                                                                                                 'v. 3.1 update
'    docSaved = mainDoc.Saved
'
'    If docSaved = False Then
'        iReply = MsgBox("Your document '" & mainDoc & "' contains unsaved changes." & vbNewLine & vbNewLine & _
'            "Click OK to save your document and run the macro." & vbNewLine & vbNewLine & "Click 'Cancel' to exit.", _
'                vbOKCancel, "Error 1")
'        If iReply = vbOK Then
'            CheckSave = False
'            mainDoc.Save
'        Else
'            CheckSave = True
'            Exit Function
'        End If
'    End If
'
'    '-----test protection
'    If ActiveDocument.ProtectionType <> wdNoProtection Then
'        MsgBox "Uh oh ... protection is enabled on document '" & mainDoc & "'." & vbNewLine & _
'            "Please unprotect the document and run the macro again." & vbNewLine & vbNewLine & _
'            "TIP: If you don't know the protection password, try pasting contents of this file into " & _
'            "a new file, and run the macro on that.", , "Error 2"
'        CheckSave = True
'        Exit Function
'    Else
'        CheckSave = False
'    End If

End Function

Function EndnotesExist() As Boolean
' Started from http://vbarevisited.blogspot.com/2014/03/how-to-detect-footnote-and-endnote.html
    Dim StoryRange As Range
    
    EndnotesExist = False
    
    For Each StoryRange In ActiveDocument.StoryRanges
        If StoryRange.StoryType = wdEndnotesStory Then
            EndnotesExist = True
            Exit For
        End If
    Next StoryRange
End Function

Function FootnotesExist() As Boolean
' Started from http://vbarevisited.blogspot.com/2014/03/how-to-detect-footnote-and-endnote.html
    Dim StoryRange As Range
    
    FootnotesExist = False
    
    For Each StoryRange In ActiveDocument.StoryRanges
        If StoryRange.StoryType = wdFootnotesStory Then
            FootnotesExist = True
            Exit For
        End If
    Next StoryRange
    
End Function


Function IsArrayEmpty(Arr As Variant) As Boolean


    Dim LB As Long
    Dim UB As Long
    
    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
        Exit Function
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
  On Error GoTo IsArrayEmptyError
        LB = LBound(Arr)
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If
  Exit Function
IsArrayEmptyError:
  Err.Source = strModule & "IsArrayEmpty"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call GlobalCleanup
  End If
End Function


Sub CreateTextFile(strText As String, suffix As String)

    Application.ScreenUpdating = False
    
    'Create report file
    Dim activeRng As Range
    Dim activeDoc As Document
    Set activeDoc = ActiveDocument
    Set activeRng = ActiveDocument.Range
    Dim activeDocName As String
    Dim activeDocPath As String
    Dim reqReportDoc As String
    Dim reqReportDocAlt As String
    Dim fnum As Integer
    Dim TheOS As String
    Dim strMacTmp As String
    TheOS = System.OperatingSystem
    
    'activeDocName below works for .doc and .docx
    activeDocName = Left(activeDoc.Name, InStrRev(activeDoc.Name, ".do") - 1)
    activeDocPath = Replace(activeDoc.Path, activeDoc.Name, "")
    
    'create text file
    reqReportDoc = activeDocPath & activeDocName & "_" & suffix & ".txt"
    
    ''''for 32 char Mc OS bug- could check if this is Mac OS too < PART 1
    If Not TheOS Like "*Mac*" Then                      'If Len(activeDocName) > 18 Then        (legacy, does not take path into account)
        reqReportDoc = activeDocPath & "\" & activeDocName & "_" & suffix & ".txt"
    Else
        Dim placeholdDocName As String
        placeholdDocName = "filenamePlacehold_Report.txt"
        reqReportDocAlt = reqReportDoc
        strMacTmp = MacScript("path to temporary items as string")
        reqReportDoc = strMacTmp & placeholdDocName
    End If
    '''end ''''for 32 char Mc OS bug part 1
    
    'set and open file for output
    Dim E As Integer
    fnum = FreeFile()
    Open reqReportDoc For Output As fnum
    
        Print #fnum, strText

    Close #fnum
    
    ''''for 32 char Mc OS bug-<PART 2
    If reqReportDocAlt <> "" Then
    Name reqReportDoc As reqReportDocAlt
    End If
    ''''END for 32 char Mac OS bug-<PART 2
    
    '----------------open Report for user once it is complete--------------------------.
    Dim Shex As Object
    
    If Not TheOS Like "*Mac*" Then
       Set Shex = CreateObject("Shell.Application")
       Shex.Open (reqReportDoc)
    Else
        MacScript ("tell application ""TextEdit"" " & vbCr & _
        "open " & """" & reqReportDocAlt & """" & " as alias" & vbCr & _
        "activate" & vbCr & _
        "end tell" & vbCr)
    End If
End Sub

Function GetText(StyleName As String, Optional ReturnArray As Boolean = False) _
  As Variant
  On Error GoTo GetTextError
  
  Dim fCount As Integer
  Dim styleArray() As Variant

  fCount = 0
  
  'Move selection to start of document
  Selection.HomeKey Unit:=wdStory

      genUtils.zz_clearFind
      With Selection.Find
          .Text = ""
          .Replacement.Text = ""
          .Forward = True
          .Wrap = wdFindStop
          .Format = True
          .Style = activeDoc.Styles(StyleName)
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
      End With
  
  Do While Selection.Find.Execute = True And fCount < 100            'fCount < 100 so we don't get an infinite loop
      fCount = fCount + 1
      
      'If paragraph return exists in selection, don't select last character (the last paragraph retunr)
      If InStr(Selection.Text, Chr(13)) > 0 Then
          Selection.MoveEnd Unit:=wdCharacter, Count:=-1
      End If
      
      'Assign selected text to variable
      ReDim Preserve styleArray(1 To fCount)
      styleArray(fCount) = Selection.Text
      
      'If the next character is a paragraph return, add that to the selection
      'Otherwise the next Find will just select the same text with the paragraph return
      If InStr(StyleName, "span") = 0 Then        'Don't select terminal para mark if char style, sends into an infinite loop
          Selection.MoveEndWhile Cset:=Chr(13), Count:=1
      End If
  Loop
      
  If fCount = 0 Then
      ReDim styleArray(1 To 1)
      styleArray(1) = ""
  End If
  
  If ReturnArray = False Then
    GetText = genUtils.GeneralHelpers.Reduce(styleArray)
  Else
    GetText = styleArray
  End If
  Exit Function
  
GetTextError:
  Err.Source = strModule & "GetText"
  If Err.Number = 5941 Or Err.Number = 5834 Then   ' The style is not present in the document
      GetText = ""
  Else
    If ErrorChecker(Err) = False Then
      Resume
    Else
      Call genUtils.GlobalCleanup
    End If
  End If
  
End Function

' ===== Reduce ================================================================
' Iterates through item passed to it (currently only an Array, but in future
' add support for Dictionary or Collection) and returns a string of all of the
' elements. Add handling in future to return other summaries (add all numbers?)

Public Function Reduce(StartGroup As Variant, Optional Delimiter As String = _
  vbNewLine) As Variant
  On Error GoTo ReduceError
  If VBA.IsArray(StartGroup) = True Then
    Dim strReturn As String
    Dim A As Long
    
    For A = LBound(StartGroup) To UBound(StartGroup)
      strReturn = strReturn & StartGroup(A)
      If A < UBound(StartGroup) Then
        strReturn = strReturn & Delimiter
      End If
    Next A
    
  Else
    ' Error if not passed an array
    Err.Raise MacError.err_NotArray
  End If
  Reduce = strReturn

  Exit Function
ReduceError:
  Err.Source = strModule & "Reduce"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function


' ===== StyleReplace ==========================================================
' Replace all instances of a specific paragraph style with a different style.
' If you want to "remove" the style, replace with "Normal" or whatever. Returns
' False if no replacements were made.

Public Function StyleReplace(SearchStyle As String, ReplaceStyle As String) As _
  Boolean
  On Error GoTo StyleReplaceError
  
  genUtils.zz_clearFind
  With activeDoc.Range.Find
    .Format = True
    .Style = SearchStyle
    .Replacement.Style = ReplaceStyle
    .Execute Replace:=wdReplaceAll
    
    If .Found = True Then
      StyleReplace = True
    Else
      StyleReplace = False
    End If
  
  End With
  genUtils.zz_clearFind
  Exit Function
  
StyleReplaceError:
  Err.Source = strModule & "StyleReplace"
  If ErrorChecker(Err, ReplaceStyle) = False Then
    Resume
  Else
    Call genUtils.GlobalCleanup
  End If
End Function
Function LoadCSVtoArray(Path As String, RemoveHeaderRow As Boolean, RemoveHeaderCol As Boolean) As Variant

'------Load CSV into 2d array, NOTE!!: base 0---------
' But also note that this now removes the header row and column too
    Dim fnum As Integer
    Dim whole_file As String
    Dim lines As Variant
    Dim one_line As Variant
    Dim num_rows As Long
    Dim num_cols As Long
    Dim the_array() As Variant
    Dim R As Long
    Dim C As Long
    
        If IsItThere(Path) = False Then
            MsgBox "There was a problem with your Castoff.", vbCritical, "Error: CSV not available"
            Exit Function
        End If
        'DebugPrint Path
        
        ' Do we need to remove a header row?
        Dim lngHeaderRow As Long
        If RemoveHeaderRow = True Then
            lngHeaderRow = 1
        Else
            lngHeaderRow = 0
        End If
        
        ' Do we need to remove a header column?
        Dim lngHeaderCol As Long
        If RemoveHeaderCol = True Then
            lngHeaderCol = 1
        Else
            lngHeaderCol = 0
        End If
        
        ' Load the csv file.
        fnum = FreeFile
        Open Path For Input As fnum
        whole_file = Input$(LOF(fnum), #fnum)
        Close fnum

        ' Break the file into lines (trying to capture whichever line break is used)
        If InStr(1, whole_file, vbCrLf) <> 0 Then
            lines = Split(whole_file, vbCrLf)
        ElseIf InStr(1, whole_file, vbCr) <> 0 Then
            lines = Split(whole_file, vbCr)
        ElseIf InStr(1, whole_file, vbLf) <> 0 Then
            lines = Split(whole_file, vbLf)
        Else
            MsgBox "There was an error with your castoff.", vbCritical, "Error parsing CSV file"
        End If

        ' Dimension the array.
        num_rows = UBound(lines)
        one_line = Split(lines(0), ",")
        num_cols = UBound(one_line)
        ReDim the_array(num_rows - lngHeaderRow, num_cols - lngHeaderCol) ' -1 if we are not using header row/col
        
        ' Copy the data into the array.
        For R = lngHeaderRow To num_rows           ' start at 1 (not 0) if we are not using the header row
            If Len(lines(R)) > 0 Then
                one_line = Split(lines(R), ",")
                For C = lngHeaderCol To num_cols   ' start at 1 (not 0) if we are not using the header column
                    'DebugPrint one_line(c)
                    the_array((R - lngHeaderRow), (C - lngHeaderCol)) = one_line(C)   ' -1 because if are not using header row/column from CSV
                Next C
            End If
        Next R
    
        ' Prove we have the data loaded.
'         DebugPrint LBound(the_array)
'         DebugPrint UBound(the_array)
'         For R = 0 To (num_rows - 1)          ' -1 again if we removed the header row
'             For c = 0 To num_cols      ' -1 again if we removed the header column
'                 DebugPrint the_array(R, c) & " | ";
'             Next c
'             DebugPrint
'         Next R
'         DebugPrint "======="
    
    LoadCSVtoArray = the_array
 
End Function

Public Function CloseOpenDocs() As Boolean
  On Error GoTo CloseOpenDocsError
    '-------------Check for/close open documents-------------------------------
    Dim strInstallerName As String
    Dim strSaveWarning As String
    Dim objDocument As Document
    Dim B As Long
    Dim Doc As Document
    
    strInstallerName = ThisDocument.Name

    If Documents.Count > 1 Then
      strSaveWarning = "All other Word documents must be closed to run the macro." & vbNewLine & vbNewLine & _
        "Click OK and I will save and close your documents." & vbNewLine & _
        "Click Cancel to exit without running the macro and close the documents yourself."
      If MsgBox(strSaveWarning, vbOKCancel, "Close documents?") = vbCancel Then
          ActiveDocument.Close
          Exit Function
      Else
        For Each Doc In Documents
            'DebugPrint doc.Name
          'But don't close THIS document
          If Doc.Name <> strInstallerName Then
              'separate step to trigger Save As prompt for previously unsaved docs
              Doc.Save
              Doc.Close
          End If
        Next Doc
      End If
    End If
  Exit Function
  
CloseOpenDocsError:
  Err.Source = strModule & "CloseOpenDocs"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GlobalCleanup
  End If
End Function





Function StartupSettings(Optional StoriesUsed As Variant, Optional AcceptAll As Boolean = False) As Boolean
  On Error GoTo StartupSettingsError
' records/adjusts/checks settings and stuff before running the rest of the macro
' returns TRUE if some check is bad and we can't run the macro
  StartupSettings = False
' mainDoc will only do stuff to main body text, not EN or FN stories. So
' do all main-text-only stuff first, then loop through stories
  Dim mainDoc As Document
  Set mainDoc = activeDoc
  
' Section of registry/preferences file to store settings
  Dim strSection As String
  strSection = "MACMILLAN_MACROS"
  
' check if file has been saved
  Dim iReply As Integer
  
'  Dim docSaved As Boolean
'  docSaved = mainDoc.Saved

' Commenting out MsgBox for now, return after writing wrapper function for
' Bookmaker validator
'  If docSaved = False Then
'        iReply = MsgBox("Your document '" & mainDoc & "' contains unsaved changes." & vbNewLine & vbNewLine & _
'            "Click OK to save your document and run the macro." & vbNewLine & vbNewLine & "Click 'Cancel' to exit.", _
'                vbOKCancel, "Error 1")
'        If iReply = vbOK Then
'    mainDoc.Save
'        Else
'            StartupSettings = True
'            Exit Function
'        End If
'  End If
    
    
' check if file has doc protection on, prompt user and quit function if it does
' For now Err.Raise should divert the code from the MsgBox call, But I'll rm anyway
  If mainDoc.ProtectionType <> wdNoProtection Then
    Err.Raise MacError.err_DocProtectionOn
'    MsgBox "Uh oh ... protection is enabled on document '" & mainDoc & "'." & vbNewLine & _
'      "Please unprotect the document and run the macro again." & vbNewLine & vbNewLine & _
'      "TIP: If you don't know the protection password, try pasting contents of this file into " & _
'      "a new file, and run the macro on that.", , "Error 2"
    StartupSettings = True
    Exit Function
  Else
    StartupSettings = False
  End If
  
  ' ========== Turn off screen updating ==========
  Application.ScreenUpdating = False
  
  ' ========== STATUS BAR: store current setting and display ==========
  System.ProfileString(strSection, "Current_Status_Bar") = Application.DisplayStatusBar
  Application.DisplayStatusBar = True
    
  ' ========== Remove bookmarks ==========
  Dim bkm As Bookmark
  
  For Each bkm In mainDoc.Bookmarks
    bkm.Delete
  Next bkm
    
  ' ========== Save current cursor location in a bookmark ==========
  ' Store current story, so we can return to it before selecting bookmark in Cleanup
  System.ProfileString(strSection, "Current_Story") = Selection.StoryType
  ' next line required for Mac to prevent problem where original selection blinked repeatedly when reselected at end
  Selection.Collapse Direction:=wdCollapseStart
  mainDoc.Bookmarks.Add Name:="OriginalInsertionPoint", Range:=Selection.Range
    
  ' ========== TRACK CHANGES: store current setting, turn off ==========
  ' ==========   OPTIONAL: Check if changes present and offer to accept all ==========
  System.ProfileString(strSection, "Current_Tracking") = mainDoc.TrackRevisions
  mainDoc.TrackRevisions = False
  
'  If AcceptAll = True Then
    If FixTrackChanges = False Then
        StartupSettings = True
    End If
'  End If
    
    
' ========== Delete field codes ==========
' Fields break cleanup and char styles, so we delete them (but retain their
' result, if any). Furthermore, fields make no sense in a manuscript, so
' even if they didn't break anything we don't want them.
' Note, however, that even though linked endnotes and footnotes are
' types of fields, this loop doesn't affect them.
    
  Dim A As Long
  Dim thisRange As Range
  Dim objField As Field
  Dim strContent As String
  
  ' StoriesUsed is optional; if an array of stories is not passed, just use the main text story here
  If IsArrayEmpty(StoriesUsed) = True Then
    ReDim StoriesUsed(1 To 1)
    StoriesUsed(1) = wdMainTextStory
  End If

  For A = LBound(StoriesUsed) To UBound(StoriesUsed)
    Set thisRange = activeDoc.StoryRanges(StoriesUsed(A))
    For Each objField In thisRange.Fields
'            DebugPrint thisRange.Fields.Count
      If thisRange.Fields.Count > 0 Then
        With objField
'                    DebugPrint .Index & ": " & .Kind
            ' None or Cold means it has no result, so we just delete
          If .Kind = wdFieldKindNone Or .Kind = wdFieldKindCold Then
            .Delete
          Else ' It has a result, so we replace field w/ just its content
            strContent = .result
            .Select
            .Delete
            Selection.InsertAfter strContent
          End If
        End With
      End If
    Next objField

  Next A

    
  ' ========== Remove content controls ==========
  ' Content controls also break character styles and cleanup
  ' They are used by some imprints for frontmatter templates
  ' for editorial, though.
  ' Doesn't work at all for a Mac, so...
  #If Win32 Then
      ClearContentControls
  #End If
  Exit Function
  
StartupSettingsError:
  Err.Source = strModule & "StartupSettings"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function


Private Function FixTrackChanges() As Boolean
' returns True if changes were fixed or not present, False if changes remain in doc
  On Error GoTo FixTrackChangesError
    Dim N As Long
    Dim oComments As Comments
    Set oComments = ActiveDocument.Comments
    
    FixTrackChanges = True
    
    'See if there are tracked changes or comments in document
    On Error Resume Next
    Selection.HomeKey Unit:=wdStory   'start search at beginning of doc
    WordBasic.NextChangeOrComment       'search for a tracked change or comment. error if none are found.
    
    ' Commenting out MsgBox for Bookmaker Validator right now
    ' When have MsgBox wrapper function, can turn it back on.
'    'If there are changes, ask user if they want macro to accept changes or cancel
    If Err.Number = 0 Then
'        If MsgBox("Bookmaker doesn't like comments or tracked changes, but it appears that you have some in your document." _
'            & vbCr & vbCr & "Click OK to ACCEPT ALL CHANGES and DELETE ALL COMMENTS right now and continue with the Bookmaker Requirements Check." _
'            & vbCr & vbCr & "Click CANCEL to stop the Bookmaker Requirements Check and deal with the tracked changes and comments on your own.", _
'            273, "Are those tracked changes I see?") = vbCancel Then           '273 = vbOkCancel(1) + vbCritical(16) + vbDefaultButton2(256)
'                FixTrackChanges = False
'                Exit Function
'        Else 'User clicked OK, so accept all tracked changes and delete all comments
            activeDoc.AcceptAllRevisions
            For N = oComments.Count To 1 Step -1
                oComments(N).Delete
            Next N
            Set oComments = Nothing
'        End If
    Else
      FixTrackChanges = True
    End If
    Exit Function
    
FixTrackChangesError:
  Err.Source = strModule & "FixTrackChanges"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
    
End Function


Private Sub ClearContentControls()
  On Error GoTo ClearContentControlsError
    'This is it's own sub because doesn't exist in Mac Word, breaks whole sub if included
    Dim cc As ContentControl
    
    For Each cc In ActiveDocument.ContentControls
        cc.Delete
    Next
  Exit Sub
ClearContentControlsError:
  Err.Source = strModule & "ClearContentControls"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub



Sub Cleanup()
  On Error GoTo CleanUpError
    ' resets everything from StartupSettings sub.
    Dim cleanupDoc As Document
    Set cleanupDoc = ActiveDocument
    
    ' Section of registry/preferences file to get settings from
    Dim strSection As String
    strSection = "MACMILLAN_MACROS"
    
    ' restore Status Bar to original setting
    ' If key doesn't exist, set to True as default
    Dim currentStatus As String
    currentStatus = System.ProfileString(strSection, "Current_Status_Bar")
    
    If currentStatus <> vbNullString Then
        Application.StatusBar = currentStatus
    Else
        Application.StatusBar = True
    End If
    
    ' reset original Track Changes setting
    ' If key doesn't exist, set to false as default
    Dim currentTracking As String
    currentTracking = System.ProfileString(strSection, "Current_Tracking")
    
    If currentTracking <> vbNullString Then
        cleanupDoc.TrackRevisions = currentTracking
    Else
        cleanupDoc.TrackRevisions = False
    End If
    
    ' return to original cursor position
    ' If key doesn't exist, search in main doc
    Dim currentStory As WdStoryType
    currentStory = System.ProfileString(strSection, "Current_Story")
    
    If cleanupDoc.Bookmarks.Exists("OriginalInsertionPoint") = True Then
        If currentStory = 0 Then
            cleanupDoc.StoryRanges(currentStory).Select
        Else
            cleanupDoc.StoryRanges(wdMainTextStory).Select
        End If
        
        Selection.GoTo what:=wdGoToBookmark, Name:="OriginalInsertionPoint"
        cleanupDoc.Bookmarks("OriginalInsertionPoint").Delete
    End If
    
    ' Turn Screen Updating on and refresh screen
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    Exit Sub
CleanUpError:
  Err.Source = strModule & "CleanUp"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Function IsReadOnly(Path As String) As Boolean
    ' Tests if the file or directory is read-only -- does NOT test if file exists,
    ' because sometimes you'll need to do that before this anyway to do something
    ' different.
    
    ' Mac 2011 can't deal with file paths > 32 char
    If IsOldMac() = True Then
        Dim strScript As String
        Dim blnWritable As Boolean
        
        strScript = _
            "set p to POSIX path of " & Chr(34) & Path & Chr(34) & Chr(13) & _
            "try" & Chr(13) & _
            vbTab & "do shell script " & Chr(34) & "test -w \" & Chr(34) & _
            "$(dirname " & Chr(34) & " & quoted form of p & " & Chr(34) & _
            ")\" & Chr(34) & Chr(34) & Chr(13) & _
            vbTab & "return true" & Chr(13) & _
            "on error" & Chr(13) & _
            vbTab & "return false" & Chr(13) & _
            "end try"
            
        blnWritable = MacScript(strScript)
        
        If blnWritable = True Then
            IsReadOnly = False
        Else
            IsReadOnly = True
        End If
    Else
        If (GetAttr(Path) And vbReadOnly) <> 0 Then
            IsReadOnly = True
        Else
            IsReadOnly = False
        End If
    End If

IsReadOnlyFinish:
    Exit Function

IsReadOnlyError:
    Err.Source = strModule & "IsReadOnly"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        genUtils.GeneralHelpers.GlobalCleanup
    End If
End Function


Public Function ReadTextFile(Path As String, Optional FirstLineOnly As Boolean = True) As String
  On Error GoTo ReadTextFileError
' load string from text file

    Dim fnum As Long
    Dim strTextWeWant As String
    
    fnum = FreeFile()
    Open Path For Input As fnum
    
    If FirstLineOnly = False Then
        strTextWeWant = Input$(LOF(fnum), #fnum)
    Else
        Line Input #fnum, strTextWeWant
    End If
    
    Close fnum
    
    ReadTextFile = strTextWeWant
  Exit Function
  
ReadTextFileError:
    Err.Source = strModule & "ReadTextFile"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        genUtils.GeneralHelpers.GlobalCleanup
    End If
End Function


Function HiddenTextSucks(StoryType As WdStoryType) As Boolean
'v. 3.1 patch : redid this whole thing as an array, addedsmart quotes, wrap toggle var
  On Error GoTo HiddenTextSucksError
'    DebugPrint StoryType
    Dim activeRng As Range
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    ' No, really, it does. Why is that even an option?
    ' Seriously, this just deletes all hidden text, based on the
    ' assumption that if it's hidden, you don't want it.
    ' returns a Boolean in case we want to notify user at some point
    
    HiddenTextSucks = False
    
    ' If Hidden text isn't shown, it won't be deleted, which
    ' defeats the purpose of doing this at all.
    Dim blnCurrentHiddenView As Boolean
    blnCurrentHiddenView = ActiveDocument.ActiveWindow.View.ShowAll
    ActiveDocument.ActiveWindow.View.ShowAll = True

    
    Dim aCounter As Long
    aCounter = 0
    
    ' Select whole doc (story, actually)
    activeRng.Select

    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Font.Hidden = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute ReplaceWith:="", Replace:=wdReplaceAll
    End With
    
    Do While Selection.Find.Execute = True And aCounter < 500
        'aCounter < 500 so we don't get an infinite loop
        aCounter = aCounter + 1
        HiddenTextSucks = True
    Loop
    
    ' Now restore Hidden Text view settings
    activeDoc.ActiveWindow.View.ShowAll = blnCurrentHiddenView
    Exit Function
HiddenTextSucksError:
    Err.Source = strModule & "HiddenTextSucks"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        genUtils.GeneralHelpers.GlobalCleanup
    End If
    
End Function


Sub ClearPilcrowFormat(StoryType As WdStoryType)
 On Error GoTo ClearPilcrowFormatError
' A pilcrow is the paragraph mark symbol. This clears all formatting and styles from
' pilcrows as found via ^p
    ' Change to story ranges?
    Dim activeRange As Range
    Set activeRange = ActiveDocument.StoryRanges(StoryType)

    With activeRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^13"       ' need to use ^13 if using wildcards
        .Replacement.Text = "^p"    ' DON'T replace with ^13, removes para style
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Replacement.Style = "Default Paragraph Font"
        .Replacement.Font.Italic = False
        .Replacement.Font.Bold = False
        .Replacement.Font.Underline = wdUnderlineNone
        .Replacement.Font.AllCaps = False
        .Replacement.Font.SmallCaps = False
        .Replacement.Font.StrikeThrough = False
        .Replacement.Font.Subscript = False
        .Replacement.Font.Superscript = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
  Exit Sub
ClearPilcrowFormatError:
  Err.Source = strModule & "ClearPilcrowFormat"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Sub StyleAllHyperlinks(Optional StoriesInUse As Variant)
  On Error GoTo StyleAllHyperlinksError
    ' StoriesInUse is an array of wdStoryTypes in use
    ' Clears active links and adds macmillan URL char styles
    ' to any proper URLs.
    ' Breaking up into sections because AutoFormat does not apply hyperlinks to FN/EN stories.
    ' Also if you AutoFormat a second time it undoes all of the formatting already applied to hyperlinks
    
    Dim S As Long
    
    Call zz_clearFind
    
    For S = LBound(StoriesInUse) To UBound(StoriesInUse)
        'Styles hyperlinks, must be performed after PreserveWhiteSpaceinBrkStylesA
        Call StyleHyperlinksA(StoryType:=(StoriesInUse(S)))
    Next S
    
    Call AutoFormatHyperlinks
    
    For S = LBound(StoriesInUse) To UBound(StoriesInUse)
        Call StyleHyperlinksB(StoryType:=(StoriesInUse(S)))
    Next S
 Exit Sub
StyleAllHyperlinksError:
  Err.Source = strModule & "StyleAllHyperlinks"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Private Sub StyleHyperlinksA(StoryType As WdStoryType)
 On Error GoTo StyleHyperlinksAError
    ' PRIVATE, if you want to style hyperlinks from another module,
    ' call StyleAllHyperlinks sub above.
    ' added by Erica 2014-10-07, v. 3.4
    ' removes all live hyperlinks but leaves hyperlink text intact
    ' then styles all URLs as "span hyperlink (url)" style
    ' -----------------------------------------
    ' this first bit removes all live hyperlinks from document
    ' we want to remove these from urls AND text; will add back to just urls later
    Dim activeRng As Range

    Set activeRng = activeDoc.StoryRanges(StoryType)
    ' remove all embedded hyperlinks regardless of character style
    ' Must use Fields obj not Hyperlink obj because if "empty" hyperlink is in
    ' doc, it will return as part of the Hyperlinks collection but will error
    ' when try to delete or access any properties.
    Dim fld As Field
    If activeRng.Fields.Count > 0 Then
      For Each fld In activeRng.Fields
      ' wdFieldKindNone = invalid field
        If fld.Kind <> wdFieldKindNone And fld.Type = wdFieldHyperlink Then
        ' If field is a link but no text appears in the document for it,
        ' just delete the whole thing (otherwise replace link w/ display text)
          If Len(fld.result.Text) = 0 Then
            fld.Delete
          Else
            fld.Unlink
          End If
        End If
      Next fld
    End If

    '------------------------------------------
    'removes all hyperlink styles
    Dim HyperlinkStyleArray(3) As String
    Dim P As Long
    
    HyperlinkStyleArray(1) = "Hyperlink"        'built-in style applied automatically to links
    HyperlinkStyleArray(2) = "FollowedHyperlink"    'built-in style applied automatically
    HyperlinkStyleArray(3) = "span hyperlink (url)" 'Macmillan template style for links
    
    For P = 1 To UBound(HyperlinkStyleArray())
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Style = HyperlinkStyleArray(P)
            .Replacement.Style = ActiveDocument.Styles("Default Paragraph Font")
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next
    
    Exit Sub
    
StyleHyperlinksAError:
        '5834 means item does not exist
        '5941 means style not present in collection
        If Err.Number = 5834 Or Err.Number = 5941 Then
            
            'If style is not present, add style
            Dim myStyle As Style
            Set myStyle = ActiveDocument.Styles.Add(Name:="span hyperlink (url)", Type:=wdStyleTypeCharacter)
            Resume
'            ' Used to add highlight color, but actually if style is missing, it's
'            ' probably a MS w/o Macmillan's styles and the highlight will be annoying.
'            'If missing style was Macmillan built-in style, add character highlighting
'            If myStyle = "span hyperlink (url)" Then
'                ActiveDocument.Styles("span hyperlink (url)").Font.Shading.BackgroundPatternColor = wdColorPaleBlue
'            End If
        Else
          Err.Source = strModule & "StyleHyperlinksA"
          If genUtils.GeneralHelpers.ErrorChecker(Err) = False Then
            Resume
          Else
            Call genUtils.GeneralHelpers.GlobalCleanup
          End If
        End If

End Sub

Private Sub AutoFormatHyperlinks()
  On Error GoTo AutoFormatHyperlinksError
    ' PRIVATE, if you want to style hyperlinks from another module,
    ' call StyleAllHyperlinks sub above.
    '--------------------------------------------------
    ' converts all URLs to hyperlinks with built-in "Hyperlink" style
    ' because some show up as plain text
    ' Note this also removes all blank paragraphs regardless of style,
    ' so needs to come after sub PreserveWhiteSpaceinBrkA
    
    
    Dim f1 As Boolean, f2 As Boolean, f3 As Boolean
    Dim f4 As Boolean, f5 As Boolean, f6 As Boolean
    Dim f7 As Boolean, f8 As Boolean, f9 As Boolean
    Dim f10 As Boolean
      
    'This first bit autoformats hyperlinks in main text story
    With Options
        ' Save current AutoFormat settings
        f1 = .AutoFormatApplyHeadings
        f2 = .AutoFormatApplyLists
        f3 = .AutoFormatApplyBulletedLists
        f4 = .AutoFormatApplyOtherParas
        f5 = .AutoFormatReplaceQuotes
        f6 = .AutoFormatReplaceSymbols
        f7 = .AutoFormatReplaceOrdinals
        f8 = .AutoFormatReplaceFractions
        f9 = .AutoFormatReplacePlainTextEmphasis
        f10 = .AutoFormatReplaceHyperlinks
        ' Only convert URLs
        .AutoFormatApplyHeadings = False
        .AutoFormatApplyLists = False
        .AutoFormatApplyBulletedLists = False
        .AutoFormatApplyOtherParas = False
        .AutoFormatReplaceQuotes = False
        .AutoFormatReplaceSymbols = False
        .AutoFormatReplaceOrdinals = False
        .AutoFormatReplaceFractions = False
        .AutoFormatReplacePlainTextEmphasis = False
        .AutoFormatReplaceHyperlinks = True
        ' Perform AutoFormat
        activeDoc.Content.AutoFormat
        ' Restore original AutoFormat settings
        .AutoFormatApplyHeadings = f1
        .AutoFormatApplyLists = f2
        .AutoFormatApplyBulletedLists = f3
        .AutoFormatApplyOtherParas = f4
        .AutoFormatReplaceQuotes = f5
        .AutoFormatReplaceSymbols = f6
        .AutoFormatReplaceOrdinals = f7
        .AutoFormatReplaceFractions = f8
        .AutoFormatReplacePlainTextEmphasis = f9
        .AutoFormatReplaceHyperlinks = f10
    End With
    
    'This bit autoformats hyperlinks in endnotes and footnotes
    ' from http://www.vbaexpress.com/forum/showthread.php?52466-applying-hyperlink-styles-in-footnotes-and-endnotes
    Dim oTemp As Document
    Dim oNote As Range
    Dim oRng As Range

    Set oTemp = Documents.Add(Visible:=False)
    
    If activeDoc.Footnotes.Count >= 1 Then
        Dim oFN As Footnote
        For Each oFN In activeDoc.Footnotes
            Set oNote = oFN.Range
            Set oRng = oTemp.Range
            oRng.FormattedText = oNote.FormattedText
            'oRng.Style= "Footnote Text"
            Options.AutoFormatReplaceHyperlinks = True
            oRng.AutoFormat
            oRng.End = oRng.End - 1
            oNote.FormattedText = oRng.FormattedText
        Next oFN
        Set oFN = Nothing
    End If
    
    If activeDoc.Endnotes.Count >= 1 Then
        Dim oEN As Endnote
        For Each oEN In activeDoc.Endnotes
            Set oNote = oEN.Range
            Set oRng = oTemp.Range
            oRng.FormattedText = oNote.FormattedText
            'oRng.Style= "Endnote Text"
            Options.AutoFormatReplaceHyperlinks = True
            oRng.AutoFormat
            oRng.End = oRng.End - 1
            oNote.FormattedText = oRng.FormattedText
        Next oEN
        Set oEN = Nothing
    End If
    
    oTemp.Close SaveChanges:=wdDoNotSaveChanges
    Set oTemp = Nothing
    Set oRng = Nothing
    Set oNote = Nothing
  Exit Sub
  
AutoFormatHyperlinksError:
  Err.Source = strModule & "AutoFormatHyperlinks"
  If genUtils.GeneralHelpers.ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Private Sub StyleHyperlinksB(StoryType As WdStoryType)
  On Error GoTo StyleHyperlinksBError
    ' PRIVATE, if you want to style hyperlinks from another module,
    ' call StyleAllHyperlinks sub above.
    '--------------------------------------------------
    ' apply macmillan URL style to hyperlinks we just tagged in Autoformat
    Dim activeRng As Range
    Set activeRng = activeDoc.StoryRanges(StoryType)
    With activeRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = "Hyperlink"
        .Replacement.Style = ActiveDocument.Styles("span hyperlink (url)")
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' -----------------------------------------------
    ' Removes all hyperlinks from the document (that were added with AutoFormat)
    ' Text to display is left intact, macmillan style is left intact
    With activeRng
        While .Hyperlinks.Count > 0
            .Hyperlinks(1).Delete
        Wend
    End With
  Exit Sub
StyleHyperlinksBError:
  Err.Source = strModule & "StyleHyperlinksB"
  If genUtils.GeneralHelpers.ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub


' ===== IsNewLine =============================================================
' Returns True if the string parameter contains a new line and ONLY a new line.
' Making a separate function to test all kinds of new lines, since sometimes
' files contain mixed line ending characters. Also note that vbNewLine constant
' returns different on Win vs. Mac 2011 vs. Mac 2016

Public Function IsNewLine(strValue As String) As Boolean
  On Error GoTo IsNewLineError
  If strValue = vbCr Or strValue = vbLf Or strValue = vbCr & vbLf Then
    IsNewLine = True
  End If
  Exit Function
  
IsNewLineError:
  Err.Source = strModule & "IsNewLine"
  If genUtils.GeneralHelpers.ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function
