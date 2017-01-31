Attribute VB_Name = "ClassHelpers"
' =============================================================================
'     CLASS HELPERS
' =============================================================================
' By Erica Warren - erica.warren@macmillan.com
'
' ===== USE ===================================================================
' Some procedures needed for classes can't actually exist in the class for one
' reason or another, so we put them here.
'
' ===== DEPENDENCIES ==========================================================
' Obviously, the class in question must be a module in the same project.


Option Explicit
Private Const strClassHelpers As String = "genUtils.ClassHelpers."

' =============================================================================
'       INSTANCING
' =============================================================================
' To be used outside of their project, custom class modules must have their
' `Instancing` property set to `PublicNotCreateable` -- so they can be used but
' an outside project can't instantiate a new object of that class. So this is
' a set of functions to call that will create a new instance in the project
' and return it to the other project to use.

' ===== NewDictionary =========================================================
Public Function NewDictionary() As genUtils.Dictionary
  Set NewDictionary = New genUtils.Dictionary
End Function


' =============================================================================
'     JSON HELPERS
' =============================================================================
' Not technically a class, but I don't feel like forking that repo now to add
' these additional functions, so I'll drop them here.

' ===== ReadJson ==============================================================
' To get from JSON file to Dictionary object, must read file to string, then
' convert string to Dictionary. This does all of that (and some error handling)

Public Function ReadJson(JsonPath As String) As genUtils.Dictionary
  On Error GoTo ReadJsonError
  Dim dictJson As genUtils.Dictionary
  
  If IsItThere(JsonPath) = True Then
    Dim strJson As String
    
    strJson = GeneralHelpers.ReadTextFile(JsonPath, False)
    If strJson <> vbNullString Then
      Set dictJson = genUtils.JsonConverter.ParseJson(strJson)
    Else
      ' If file exists but has no content, return empty dictionary
      Set dictJson = ClassHelpers.NewDictionary
    End If
  Else
    Err.Raise MacError.err_FileNotThere
  End If
  
  If dictJson Is Nothing Then
    DebugPrint "ReadJson fail"
  Else
'    DebugPrint "dictJson count: " & dictJson.Count
  End If
  
  Set ReadJson = dictJson
  Exit Function
  
ReadJsonError:
  Err.Source = strClassHelpers & "ReadJson"
  If ErrorChecker(Err, JsonPath) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function


' ===== WriteJson =============================================================
' JsonConverter.ConvertToJson returns a string, when we then need to write to
' a text file if we want the output. This combines those. Will overwrite the
' original file if already exists, will create file if it does not.

Public Sub WriteJson(JsonPath As String, JsonData As genUtils.Dictionary)
  On Error GoTo WriteJsonError:

  Dim strJson As String
  strJson = JsonConverter.ConvertToJson(JsonData, Whitespace:=2)
  ' `OverwriteTextFile` validates directory
  GeneralHelpers.OverwriteTextFile JsonPath, strJson
  Exit Sub
  
WriteJsonError:
  Err.Source = strClassHelpers & "WriteJson"
  If ErrorChecker(Err, JsonPath) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub


' ===== AddToJson =============================================================
' Adds the key/value pair to an already existing JSON file. Creates file if it
' doesn't exist yet. `NewValue` can be anything valid for JSON: string,
' number, boolean, dictionary, array. `JsonFile` is full path to file.

' NOTE!! If `NewKey` already exists, the value will be overwritten. Could change
' to check for existance and do something else instead (append number to key,
' add value to array, return false, whatever).

Public Sub AddToJson(JsonFile As String, NewKey As String, NewValue As Variant)
  On Error GoTo AddToJsonError
  Dim dictJson As genUtils.Dictionary
  
  ' READ JSON FILE IF IT EXISTS
  ' Does the file exist yet?
  If GeneralHelpers.IsItThere(JsonFile) = True Then
    Set dictJson = ReadJson(JsonFile)
  Else
    ' File doesn't exist yet, we'll be creating it
    Set dictJson = New genUtils.Dictionary
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

  Exit Sub

AddToJsonError:
  Err.Source = strClassHelpers & "AddToJson"
  If ErrorChecker(Err, JsonFile) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

' =============================================================================
'       DICTIONARY HELPERS
' =============================================================================

' ===== MergeDictionary =======================================================
' Add all key:value pairs of one dictionary to another dictionary. Default is
' to overwrite value in DictOne if a key in DictTwo matches; Overwrite = False
' adds an integer to the key name and adds a new key:value pair.

Public Function MergeDictionary(DictOne As genUtils.Dictionary, DictTwo As _
  genUtils.Dictionary, Optional Overwrite As Boolean = True) As _
  genUtils.Dictionary
  On Error GoTo MergeDictionaryError
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
  Exit Function

MergeDictionaryError:
  Err.Source = strClassHelpers & "MergeDictionary"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function

' =============================================================================
'       PROGRESS BAR HELPERS
' =============================================================================
'
' ===== UpdateBarAndWait ======================================================
' Call this to update the progress bar. Can't be in the ProgressBar class,
' because that class can crash the program if it gets called
' again before it finishes the first call. This includes `DoEvents` which allows
' other work to finish before returning to the calling procedure.

Public Sub UpdateBarAndWait(Bar As ProgressBar, Status As String, _
  Percent As Single)
    On Error GoTo UpdateBarAndWaitError
    Bar.Done = False
    Bar.Increment Percent, Status
    Do
        DoEvents  ' Allows other macro execution to continue
    Loop Until Bar.Done = True
    DebugPrint "Progress: " & (Percent * 100) & "%"
  Exit Sub

UpdateBarAndWaitError:
  Err.Source = strClassHelpers & "UpdateBarAndWaitError"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

