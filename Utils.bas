Attribute VB_Name = "Utils"
' =============================================================================
'         UTILS
' =============================================================================

' ***** USE *******************************************************************
' Procedures that can be used in multiple VBA projects. Trying to be as general
' purpose as possible.

' ***** DEPENDENCIES **********************************************************
' Must be imported with other submodules: VBA-Dictionary, VBA-JSON


' *****************************************************************************
'     DECLARATIONS
' *****************************************************************************
Option Explicit

' assign to actual document we're working on
' TODO: probably better managed via a class
Public activeDoc As Document


' *****************************************************************************
'     PROCEDURES
' *****************************************************************************

' ===== GetFileExtension =======================================================
' Returns file extension WITHOUT dot

Public Function GetFileExtension(File As String) As String
  GetFileExtension = Right(File, InStr(StrReverse(File), ".") - 1)
End Function

' ===== GetFileNameOnly ========================================================
' Strips file extension from end and path from beginning of string.

Public Function GetFileNameOnly(File As String) As String
  Dim lngLastSeparatorPosition As Long
  Dim lngExtensionDotLen As Long
  Dim lngFileNameStart As Long
  Dim lngFileNameLength As Long
  
  lngLastSeparatorPosition = InStrRev(File, Application.PathSeparator)
  lngExtensionDotLen = InStr(StrReverse(File), ".")
  lngFileNameStart = lngLastSeparatorPosition + 1
  lngFileNameLength = Len(File) - lngLastSeparatorPosition - lngExtensionDotLen
  GetFileNameOnly = Mid(File, lngFileNameStart, lngFileNameLength)
End Function

' ===== GetFileName ===========================================================
' Strips path and returns file name with extension.

Public Function GetFileName(File As String) As String
  GetFileName = Right(File, InStr(StrReverse(File), "."))
End Function


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

' ===== IsOldMac ==============================================================
' Checks this is a Mac running Office 2011 or earlier. Good for things like
' checking if we need to account for file paths > 3 char (which 2011 can't
' handle but Mac 2016 can.

Public Function IsOldMac() As Boolean
  IsOldMac = False
  #If Mac Then
      If Application.Version < 16 Then
          IsOldMac = True
      End If
  #End If
End Function

' ===== DocPropExists =========================================================
' Tests if a particular custom document property exists in the document. If
' it's already a Document object we already know that it exists and is open
' so we don't need to test for those here. Should be tested somewhere in
' calling procedure though.

Public Function DocPropExists(objDoc As Document, PropName As String) As Boolean
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
End Function

' ===== IsOpen ================================================================
' Tests if the Word document is currently open.

Public Function IsOpen(DocPath As String) As Boolean
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
    End If
  End If
End Function

' ===== IsWordFormat ==========================================================
' Checks extension to see if file is a Word document or template. Notably,
' does not test if it's a file type that Word CAN open (e.g., .html), just
' if it's a native Word file type.

' Ignores final character for newer file types, just checks for .dot / .doc

Public Function IsWordFormat(PathToFile As String) As Boolean
  Dim strExt As String
  strExt = Left(Right(PathToFile, InStr(StrReverse(PathToFile), ".")), 4)
  If strExt = ".dot" Or strExt = ".doc" Then
    IsWordFormat = True
  Else
    IsWordFormat = False
  End If
End Function

' ===== IsLocked ==============================================================
' Tests if any file is locked by some kind of process.

Public Function IsLocked(FilePath As String) As Boolean
  On Error GoTo IsLockedError
  IsLocked = False
  If IsItThere(FilePath) = False Then
    Exit Function
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
  If Err.Number = 70 Or Err.Number = 75 Then
      IsLocked = True
  End If
End Function

' ===== IsItThere =============================================================
' Check if file or directory exists on PC or Mac.
' Dir() doesn't work on Mac 2011 if file is longer than 32 char

Public Function IsItThere(Path As String) As Boolean
  'Remove trailing path separator from dir if it's there
  If Right(Path, 1) = Application.PathSeparator Then
    Path = Left(Path, Len(Path) - 1)
  End If

  If IsOldMac = True Then
    Dim strScript As String
    strScript = "tell application " & Chr(34) & "System Events" & Chr(34) & _
        "to return exists disk item (" & Chr(34) & Path & Chr(34) _
        & " as string)"
    IsItThere = ShellAndWaitMac(strScript)
  Else
    Dim strCheckDir As String
    strCheckDir = Dir(Path, vbDirectory)
    
    If strCheckDir = vbNullString Then
        IsItThere = False
    Else
        IsItThere = True
    End If
  End If
End Function

' ===== ParentDirExists =======================================================
' If `FilePath` is the full path to a file (that may or may not exist), then
' this checks that the directory the file is in exists. Good for checking paths
' to files before you create them.

Public Function ParentDirExists(FilePath As String) As Boolean
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
End Function

' ===== KillAll ===============================================================
' Deletes file (or folder?) on PC or Mac. Mac can't use Kill() if file name
' is longer than 32 char. Returns true if successful.
    
Public Function KillAll(Path As String) As Boolean
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
        ShellAndWaitMac (strCommand)
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
      Exit Function
  End Select
End Function

' ===== IsInstalledAddInError =================================================
' Check if the file is currently loaded as an AddIn. Because we can't delete
' it if it is loaded (though we can delete it if it's just referenced but
' not loaded).

Public Function IsInstalledAddIn(FileName As String) As Boolean
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
End Function

' ===== ShellAndWaitMac =======================================================
' Sends shell command to AppleScript on Mac (to replace missing functions!)

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

' ===== OverwriteTextFile =====================================================
' Pretty self explanatory. TextFile parameter should be full path.

Public Sub OverwriteTextFile(TextFile As String, NewText As String)
  Dim FileNum As Integer
  ' Will create file if not exist, but parent dir must exist
  If ParentDirExists(TextFile) = True Then
    FileNum = FreeFile ' next file number
    Open TextFile For Output Access Write As #FileNum
    Print #FileNum, NewText ' overwrite information in the text of the file
    Close #FileNum ' close the file
  Else
    ' directory is invalid

  End If
End Sub

' ===== AppendTextFile ========================================================
' Appends Contents string to file that already exists.

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
  End If
End Sub

' ===== SetPathSeparator ======================================================
' Replaces original path separators in string with current file system separators

Public Function SetPathSeparator(strOrigPath As String) As String

  Dim strCharactersCollection As Collection
  Dim strCharacter As String
  strCharactersCollection.Add = ":"
  strCharactersCollection.Add = "/"
  strCharactersCollection.Add = "\"
  
  For Each strCharacter In strCharactersCollection
    If InStr(strOrigPath, strCharacter) > 0 Then
      strFinalPath = VBA.Replace(strOrigPath, strOrigPath, _
        Application.PathSeparator)
    End If
  Next strCharacter
  
  SetPathSeparator = strFinalPath

End Function

' ===== CloseOpenDocs =========================================================
' Closes all open Word documents.

Public Function CloseOpenDocs() As Boolean

    '-------------Check for/close open documents-------------------------------
    Dim strInstallerName As String
    Dim strSaveWarning As String
    Dim objDocument As Document
    Dim B As Long
    Dim doc As Document
    
    strInstallerName = ThisDocument.Name

    If Documents.Count > 1 Then
      strSaveWarning = "All other Word documents must be closed to run the macro." & vbNewLine & vbNewLine & _
        "Click OK and I will save and close your documents." & vbNewLine & _
        "Click Cancel to exit without running the macro and close the documents yourself."
      If MsgBox(strSaveWarning, vbOKCancel, "Close documents?") = vbCancel Then
          ActiveDocument.Close
          Exit Function
      Else
        For Each doc In Documents
            'DebugPrint doc.Name
          'But don't close THIS document
          If doc.Name <> strInstallerName Then
              'separate step to trigger Save As prompt for previously unsaved docs
              doc.Save
              doc.Close
          End If
        Next doc
      End If
    End If

End Function

' ===== IsReadOnly ============================================================
' Tests if the file or directory is read-only -- does NOT test if file exists,
' because sometimes you'll need to do that before this anyway to do something
' different.

' Mac 2011 can't deal with file paths > 32 char
    
Function IsReadOnly(Path As String) As Boolean

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
End Function

' ===== ReadTextFile ==========================================================

Public Function ReadTextFile(Path As String, Optional FirstLineOnly As Boolean _
  = True) As String

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

End Function


' =============================================================================
'     JSON HELPERS
' =============================================================================
' I don't feel like forking that repo now to add these additional functions, so
' I'll drop them here.

' ===== ReadJson ==============================================================
' To get from JSON file to Dictionary object, must read file to string, then
' convert string to Dictionary. This does all of that (and some error handling)

Public Function ReadJson(JsonPath As String) As Dictionary
  Dim dictJson As Dictionary
  
  If Utils.IsItThere(JsonPath) = True Then
    Dim strJson As String
    
    strJson = Utils.ReadTextFile(JsonPath, False)
    If strJson <> vbNullString Then
      Set dictJson = JsonConverter.ParseJson(strJson)
    Else
      ' If file exists but has no content, return empty dictionary
      Set dictJson = New Dictionary
    End If
  End If
  
  If dictJson Is Nothing Then
    Utils.DebugPrint "ReadJson fail"
  Else
'    DebugPrint "dictJson count: " & dictJson.Count
  End If
  
  Set ReadJson = dictJson
End Function


' ===== WriteJson =============================================================
' JsonConverter.ConvertToJson returns a string, when we then need to write to
' a text file if we want the output. This combines those. Will overwrite the
' original file if already exists, will create file if it does not.

Public Sub WriteJson(JsonPath As String, JsonData As Dictionary)
  Dim strJson As String
  strJson = JsonConverter.ConvertToJson(JsonData, Whitespace:=2)
  ' `OverwriteTextFile` validates directory
  Utils.OverwriteTextFile JsonPath, strJson
End Sub


' ===== AddToJson =============================================================
' Adds the key/value pair to an already existing JSON file. Creates file if it
' doesn't exist yet. `NewValue` can be anything valid for JSON: string,
' number, boolean, dictionary, array. `JsonFile` is full path to file.

' NOTE!! If `NewKey` already exists, the value will be overwritten. Could change
' to check for existance and do something else instead (append number to key,
' add value to array, return false, whatever).

Public Sub AddToJson(JsonFile As String, NewKey As String, NewValue As Variant)
  Dim dictJson As Dictionary
  
  ' READ JSON FILE IF IT EXISTS
  ' Does the file exist yet?
  If Utils.IsItThere(JsonFile) = True Then
    Set dictJson = ReadJson(JsonFile)
  Else
    ' File doesn't exist yet, we'll be creating it
    Set dictJson = New Dictionary
  End If
  
  ' ADD NEW ITEM TO DICTIONARY
  ' `.Item("key")` method will add if key is new, overwrite if not
  If VBA.IsObject(NewValue) = True Then
    ' Need `Set` keyword for object
    Set dictJson.Item(NewKey) = NewValue
  Else
    dictJson.Item(NewKey) = NewValue
  End If

  ' WRITE UPDATED DICTIONARY (BACK) TO JSON FILE
  Call WriteJson(JsonFile, dictJson)
End Sub

' =============================================================================
'       DICTIONARY HELPERS
' =============================================================================

' ===== MergeDictionary =======================================================
' Add all key:value pairs of one dictionary to another dictionary. Default is
' to overwrite value in DictOne if a key in DictTwo matches; Overwrite = False
' adds an integer to the key name and adds a new key:value pair.

Public Function MergeDictionary(DictOne As Dictionary, DictTwo As Dictionary, _
  Optional Overwrite As Boolean = True) As Dictionary

  Dim key2 As Variant
  Dim lngCount As Long
  Dim strKey As String
  
  lngCount = 0
  
  ' Use .Item() not .Add, because .Add errors if same key is used
  For Each key2 In DictTwo.Keys
    If Overwrite = False Then
      If DictOne.Exists(key2) = True Then
        lngCount = lngCount + 1
        strKey = key2 & lngCount
      Else
        strKey = key2
      End If
    Else
      strKey = key2
    End If
    
    DictOne.Item(strKey) = DictTwo(key2)
  
  Next key2
  
  Set MergeDictionary = DictOne
End Function


