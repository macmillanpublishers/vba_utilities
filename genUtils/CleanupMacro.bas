Attribute VB_Name = "CleanupMacro"
' by Erica Warren - erica.warren@macmillan.com

' ======== PURPOSE ============================
' Performs standard typographic cleanup in current document

' ======== DEPENDENCIES =======================
' 1. Requires ProgressBar userform module

Option Explicit
Option Base 1

Private activeRng As Range
Private Const strCleanup As String = "genUtils.CleanupMacro."

Public Function MacmillanManuscriptCleanup() As genUtils.Dictionary
  On Error GoTo MacmillanManuscriptCleanupError
' Just checking if it finishes at all for now. Can convert to a more detailed
' set of tests later.
  Dim dictReturn As genUtils.Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False

  '----------Timer Start-----------------------------
  'Dim StartTime As Double
  'Dim SecondsElapsed As Double
  
  'Remember time when macro starts
  '  StartTime = Timer
  
  '------------check for endnotes and footnotes--------------------------------
  Dim stStories() As Variant
  stStories = genUtils.GeneralHelpers.StoryArray
  
  ' ======= Run startup checks ========
  ' True means a check failed (e.g., doc protection on)
  If StartupSettings(StoriesUsed:=stStories) = True Then
    Err.Raise MacError.err_MacErrGeneral
  End If
  
  ' Change to just check for backtick characters
  If zz_errorChecks = True Then
    Err.Raise MacError.err_MacErrGeneral
  End If
      
  '--------Progress Bar--------------------------------------------------------
  'Percent complete and status for progress bar (PC) and status bar (Mac)
  'Requires ProgressBar custom UserForm and Class
  Dim sglPercentComplete As Single
  Dim strStatus As String
  Dim strTitle As String
  
  'First status shown will be randomly pulled from array, for funzies
  Dim funArray() As String
  ReDim funArray(1 To 10)      'Declare bounds of array here

  funArray(1) = "* Waving magic wand..."
  funArray(2) = "* Doing a little dance..."
  funArray(3) = "* Making plans for the weekend..."
  funArray(4) = "* Setting sail for the tropics..."
  funArray(5) = "* Making a cup of tea..."
  funArray(6) = "* Beep beep boop beep..."
  funArray(7) = "* Writing the next Great American Novel..."
  funArray(8) = "* My, don't you look nice today..."
  funArray(9) = "* Having a snack..."
  funArray(10) = "* Initiating launch sequence..."
  
  Dim X As Integer
  
  'Rnd returns random number between (0,1], rest of expression is to return
  ' an integer (1,10)
  Randomize           'Sets seed for Rnd below to value of system timer
  X = Int(UBound(funArray()) * Rnd()) + 1
  
  'DebugPrint x
  
  strTitle = "Macmillan Manuscript Cleanup Macro"
  sglPercentComplete = 0.05
  strStatus = funArray(X)

  Dim oProgressCleanup As ProgressBar
' Triggers Initialize event, which also triggers Show method on PC only
  Set oProgressCleanup = New ProgressBar

  oProgressCleanup.Title = strTitle
  DebugPrint "Starting Cleanup macro"
  
' This sub calls ProgressBar.Increment and waits for it to finish before
' returning here
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressCleanup, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  '-----------Delete hidden text ----------------------------------------------
  Dim S As Long
  
  For S = 1 To UBound(stStories())
  If genUtils.GeneralHelpers.HiddenTextSucks(StoryType:=(stStories(S))) = _
    True Then
    ' Notify user maybe?
    End If
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind

' ---------- Clear formatting from paragraph marks, symbols -------------------
' Per Westchester, can cause macro to break
  
  For S = 1 To UBound(stStories())
    Call genUtils.GeneralHelpers.ClearPilcrowFormat(StoryType:=(stStories(S)))
    Call genUtils.CleanupMacro.CleanSomeSymbols(StoryTypes:=(stStories(S)))
  Next S
  
  '-----------Find/Replace with Wildcards = False------------------------------
  Call genUtils.GeneralHelpers.zz_clearFind                          'Clear find object
  
  sglPercentComplete = 0.2
  strStatus = "* Fixing quotes, unicode, section breaks..." & vbCr & strStatus

  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressCleanup, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  For S = 1 To UBound(stStories())
  ' has to be alone b/c Match Wildcards has to be disabled: Smart Quotes,
  ' Unicode (ellipse), section break
    Call genUtils.CleanupMacro.RmNonWildcardItems(StoryType:=(stStories(S)))
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind

  '-------------Tag characters styled "span preserve characters"---------------
  sglPercentComplete = 0.4
  strStatus = "* Preserving styled whitespace characters..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressCleanup, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  For S = 1 To UBound(stStories())
' tags styled page breaks, tabs
    Call genUtils.CleanupMacro.PreserveStyledCharactersA(StoryType:= _
      (stStories(S)))
  Next S
  Call genUtils.GeneralHelpers.zz_clearFind
  
  '---------------Find/Replace for rest of the typographic errors----------------------
  sglPercentComplete = 0.6
  strStatus = "* Removing unstyled whitespace, fixing ellipses and dashes..." _
    & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressCleanup, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  For S = 1 To UBound(stStories())
' v. 3.7 does NOT remove manual page breaks or multiple paragraph returns
    Call genUtils.CleanupMacro.RmWhiteSpaceB(StoryType:=(stStories(S)))
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind

  '---------------Remove tags from "span preserve characters"-------------------------
  sglPercentComplete = 0.86
  strStatus = "* Cleaning up styled whitespace..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressCleanup, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  For S = 1 To UBound(stStories())
' replaces character tags with actual character
    Call genUtils.CleanupMacro.PreserveStyledCharactersB(StoryType:= _
      (stStories(S)))
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind
  
  '---------------Convert all underlines to standard-------------------------
  sglPercentComplete = 0.87
  strStatus = "* Standardizing underline format..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressCleanup, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  For S = 1 To UBound(stStories())
    Call genUtils.CleanupMacro.FixUnderlines(StoryType:=(stStories(S)))
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind

  ' ----------- remove Shape objects
  For S = 1 To UBound(stStories())
    Call genUtils.CleanupMacro.ShapeDelete
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind
  
  '-----------------Restore original settings--------------------------------------
  sglPercentComplete = 1#
  strStatus = "* Finishing up..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressCleanup, _
    Status:=strStatus, Percent:=sglPercentComplete)

  Call genUtils.GeneralHelpers.Cleanup
  Unload oProgressCleanup
    
'    MsgBox "Hurray, the Macmillan Cleanup macro has finished running! Your manuscript looks great!"                                 'v. 3.1 patch / request  v. 3.2 made a little more fun
    
    '----------------Timer End-----------------
    'Determine how many seconds code took to run
    '  SecondsElapsed = Round(Timer - StartTime, 2)
    
    'Notify user in seconds
    '  DebugPrint "This code ran successfully in " & SecondsElapsed & " seconds"
  dictReturn.Item("pass") = True
  Set MacmillanManuscriptCleanup = dictReturn
  Exit Function

MacmillanManuscriptCleanupError:
  Err.Source = strCleanup & "MacmillanManuscriptCleanup"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function


Private Sub RmNonWildcardItems(StoryType As WdStoryType)                                             'v. 3.1 patch : redid this whole thing as an array, addedsmart quotes, wrap toggle var
  On Error GoTo RmNonWildcardItemsError
  Set activeRng = activeDoc.StoryRanges(StoryType)

' number of items in array should be declared here
  Dim noWildTagArray(3) As String
  Dim noWildReplaceArray(3) As String
  Dim C As Long
  Dim wrapToggle As String
  
  wrapToggle = "wdFindContinue"
  Application.Options.AutoFormatAsYouTypeReplaceQuotes = True
  
  
  noWildTagArray(1) = "^u8230"
' stright single quote
  noWildTagArray(2) = "^39"
' straight double quote
  noWildTagArray(3) = "^34"
  
  noWildReplaceArray(1) = " . . . "
  noWildReplaceArray(2) = "'"
  noWildReplaceArray(3) = """"

  Call genUtils.GeneralHelpers.zz_clearFind

  For C = 1 To UBound(noWildTagArray())
    If C = 3 Then wrapToggle = "wdFindStop"
    
    With activeRng.Find
      .Text = noWildTagArray(C)
      .Replacement.Text = noWildReplaceArray(C)
      .Wrap = wdFindContinue
      .Execute Replace:=wdReplaceAll
    End With
  Next
  Exit Sub
  
RmNonWildcardItemsError:
  Err.Source = strCleanup & "RmNonWildcardItems"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Private Sub PreserveStyledCharactersA(StoryType As WdStoryType)
' replaces correctly styled characters with placeholder so they don't get removed
  On Error GoTo PreserveStyledCharactersAError
  Set activeRng = activeDoc.StoryRanges(StoryType)
  
  Dim preserveCharFindArray(5) As String  ' declare number of items in array
  Dim preserveCharReplaceArray(5) As String   'delcare number of items in array
  Dim preserveCharStyle As String
  Dim M As Long
  
  preserveCharStyle = "span preserve characters (pre)"
  
  Dim keyStyle As Word.Style
  Set keyStyle = activeDoc.Styles(preserveCharStyle)
  
  preserveCharFindArray(1) = "^t" 'tabs
  preserveCharFindArray(2) = "  "  ' two spaces
  preserveCharFindArray(3) = "   "    'three spaces
  preserveCharFindArray(4) = "^l"  ' soft return
  preserveCharFindArray(5) = "- "  ' hyphen + space
  
  preserveCharReplaceArray(1) = "`E|"
  preserveCharReplaceArray(2) = "`G|"
  preserveCharReplaceArray(3) = "`J|"
  preserveCharReplaceArray(4) = "`K|"
  preserveCharReplaceArray(5) = "`HS|"
  
  For M = 1 To UBound(preserveCharFindArray())
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = preserveCharFindArray(M)
      .Replacement.Text = preserveCharReplaceArray(M)
      .Wrap = wdFindContinue
      .Format = True
      .Style = preserveCharStyle
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
  Next
  
  Exit Sub
  
PreserveStyledCharactersAError:
  Err.Source = strCleanup & "PreserveStyledCharactersA"
  If ErrorChecker(Err, preserveCharStyle) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Private Sub RmWhiteSpaceB(StoryType As WdStoryType)
 On Error GoTo RmWhiteSpaceBError
 
  Set activeRng = activeDoc.StoryRanges(StoryType)

  Dim wsFindArray(33) As String              'number of items in array should be declared here
  Dim wsReplaceArray(33) As String       'and here
  Dim i As Long

  wsFindArray(1) = ".{4,}"             '4 or more consecutive periods, into proper 4 dot ellipse
  wsFindArray(2) = "..."                  '3 consecutive periods, into 3 dot ellipse
  wsFindArray(3) = "^s"                  'non-breaking space replaced with space                                 v. 3.1 patch
  wsFindArray(4) = "([! ]). . ."          'add leading space for ellipse if not present
  wsFindArray(5) = ". . .([! ])"          'add trailing space for ellipse if not present
  wsFindArray(6) = "^t{1,}"             'tabs replace with spaces
  wsFindArray(7) = "^l{1,}"               'manual line breaks replaced with hard return
  wsFindArray(8) = " {2,}"               '2 or more spaces replaced with single space
  wsFindArray(9) = "^13 "               'paragraph, space replaced with just paragraph
  wsFindArray(10) = " ^13"               'space, paragraph replaced with just paragraph
  wsFindArray(11) = "^-"                     'optional hyphen deleted                                                    v. 3.1 patch
  wsFindArray(12) = "^~"                      'non-breaking hyphen replaced with reg hyphen               v. 3.1 patch
  wsFindArray(13) = " ^= "                    'endash w/ spaces convert to emdash (no spaces)                                v. 3.1 patch
  wsFindArray(14) = "---"                   '3 hyphens to emdash
  wsFindArray(15) = "--"                   '2 hyphens to emdash                           v. 3.7 changed from en-dash to em-dash, per usual usage.
  wsFindArray(16) = " -"                  'hyphen leading space-remove
  wsFindArray(17) = "- "                  'hyphen trailing space-remove
  wsFindArray(18) = " ^+"                  'emdash leading space-remove
  wsFindArray(19) = "^+ "                  'emdash trailing space-remove
  wsFindArray(20) = " ^="                  'endash leading space-remove
  wsFindArray(21) = "^= "                  'endash trailing space-remove
  wsFindArray(22) = "\( "                     ' remove space after open parens                                                           v. 3.3
  wsFindArray(23) = " \)"                     ' removespace before closing parens                                                       v. 3.3
  wsFindArray(24) = "\[ "                     ' removespace after opening bracket                                                    v. 3.3
  wsFindArray(25) = " \]"                    ' removespace before closing bracket                                                   v. 3.3
  wsFindArray(26) = "\{ "                     ' removespace after opening curly bracket                                          v. 3.3
  wsFindArray(27) = " \}"                     ' removespace before closing curly bracket                                         v. 3.3
  wsFindArray(28) = "$ "                      ' removespace after dollar sign                                                                v. 3.3
  wsFindArray(29) = " . . . ."                ' remove space before 4-dot ellipsis (because it's a period)       v 3.7
  wsFindArray(30) = ".."                         'replace double period with single period                v. 3.7
  wsFindArray(31) = ",,"                          'replace double commas with single comma                v. 3.7

' Test if Mac or PC because character code for closing quotes is different
' on different platforms
  #If Mac Then
    wsFindArray(32) = ". . . " & Chr(211)
    wsFindArray(33) = ". . . " & Chr(213)
  #Else
    wsFindArray(32) = ". . . " & Chr(148)
    wsFindArray(33) = ". . . " & Chr(146)
  #End If

  wsReplaceArray(1) = ". . . . "      ' v. 3.2 EW removed leading space--not needed, 1st dot is a period
  wsReplaceArray(2) = " . . . "
  wsReplaceArray(3) = " "
  wsReplaceArray(4) = "\1 . . ."
  wsReplaceArray(5) = ". . . \1"
  wsReplaceArray(6) = " "
  wsReplaceArray(7) = "^p"
  wsReplaceArray(8) = " "
  wsReplaceArray(9) = "^p"
  wsReplaceArray(10) = "^p"
  wsReplaceArray(11) = ""
  wsReplaceArray(12) = "-"
  wsReplaceArray(13) = "^+"
  wsReplaceArray(14) = "^+"
  wsReplaceArray(15) = "^+"       'v. 3.7 changed to em-dash per common usage
  wsReplaceArray(16) = "-"
  wsReplaceArray(17) = "-"
  wsReplaceArray(18) = "^+"
  wsReplaceArray(19) = "^+"
  wsReplaceArray(20) = "^="
  wsReplaceArray(21) = "^="
  wsReplaceArray(22) = "("
  wsReplaceArray(23) = ")"
  wsReplaceArray(24) = "["
  wsReplaceArray(25) = "]"
  wsReplaceArray(26) = "{"
  wsReplaceArray(27) = "}"
  wsReplaceArray(28) = "$"
  wsReplaceArray(29) = ". . . ."
  wsReplaceArray(30) = "."
  wsReplaceArray(31) = ","

  #If Mac Then
    wsReplaceArray(32) = ". . ." & Chr(211)
    wsReplaceArray(33) = ". . ." & Chr(213)
  #Else
    wsReplaceArray(32) = ". . ." & Chr(148)
    wsReplaceArray(33) = ". . ." & Chr(146)
  #End If
  
  Call genUtils.GeneralHelpers.zz_clearFind
  For i = 1 To UBound(wsFindArray())
    With activeRng.Find
      .Text = wsFindArray(i)
      .Replacement.Text = wsReplaceArray(i)
      .Wrap = wdFindContinue
      .MatchWildcards = True
      .Execute Replace:=wdReplaceAll
    End With
  Next
  Exit Sub
  
RmWhiteSpaceBError:
  Err.Source = strCleanup & "RmWhiteSpaceBError"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Private Sub PreserveStyledCharactersB(StoryType As WdStoryType)
  On Error GoTo PreserveStyledCharactersBError
  ' added by EW v. 3.2
  ' replaces placeholders with original characters
  Set activeRng = activeDoc.StoryRanges(StoryType)

  Dim preserveCharFindArray(5) As String  ' declare number of items in array
  Dim preserveCharReplaceArray(5) As String   'declare number of items in array
  Dim preserveCharStyle As String
  Dim N As Long

  preserveCharStyle = "span preserve characters (pre)"

  Dim keyStyle As Word.Style
  Set keyStyle = activeDoc.Styles(preserveCharStyle)

  preserveCharFindArray(1) = "`E|" 'tabs
  preserveCharFindArray(2) = "`G|"    ' two spaces
  preserveCharFindArray(3) = "`J|"   'three spaces
  preserveCharFindArray(4) = "`K|"   ' soft return
  preserveCharFindArray(5) = "`HS|"

  preserveCharReplaceArray(1) = "^t"
  preserveCharReplaceArray(2) = "  "
  preserveCharReplaceArray(3) = "   "
  preserveCharReplaceArray(4) = "^l"
  preserveCharReplaceArray(5) = "- "

  Call genUtils.GeneralHelpers.zz_clearFind
  For N = 1 To UBound(preserveCharFindArray())
    With activeRng.Find
      .Text = preserveCharFindArray(N)
      .Replacement.Text = preserveCharReplaceArray(N)
      .Wrap = wdFindContinue
      .Format = True
      .Style = preserveCharStyle
      .Execute Replace:=wdReplaceAll
    End With
  Next
  Exit Sub

PreserveStyledCharactersBError:
  Err.Source = strCleanup & "PreserveStyledCharactersB"
  If ErrorChecker(Err, keyStyle.NameLocal) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Private Sub FixUnderlines(StoryType As WdStoryType)
' Turns out there are like 17 different types of underlines, and we don't want
' to lose them but we really just want them to be a simple, single underline.
' There has got to be a better way to loop through an enumeration but
' I've failed to find it so far

  On Error GoTo FixUnderlinesError

  Set activeRng = activeDoc.StoryRanges(StoryType)
  
  Dim strUnderlines(1 To 16) As WdUnderline
  Dim A As Long
  
  strUnderlines(1) = wdUnderlineDash
  strUnderlines(2) = wdUnderlineDashHeavy
  strUnderlines(3) = wdUnderlineDashLong
  strUnderlines(4) = wdUnderlineDashLongHeavy
  strUnderlines(5) = wdUnderlineDotDash
  strUnderlines(6) = wdUnderlineDotDashHeavy
  strUnderlines(7) = wdUnderlineDotDotDash
  strUnderlines(8) = wdUnderlineDotDotDashHeavy
  strUnderlines(9) = wdUnderlineDotted
  strUnderlines(10) = wdUnderlineDottedHeavy
  strUnderlines(11) = wdUnderlineDouble
  strUnderlines(12) = wdUnderlineThick
  strUnderlines(13) = wdUnderlineWavy
  strUnderlines(14) = wdUnderlineWavyDouble
  strUnderlines(15) = wdUnderlineWavyHeavy
  strUnderlines(16) = wdUnderlineWords

  Call genUtils.GeneralHelpers.zz_clearFind
  For A = LBound(strUnderlines()) To UBound(strUnderlines())
    With activeDoc.Range.Find
      .Wrap = wdFindStop
      .Format = True
      .Font.Underline = strUnderlines(A)
      .Replacement.Font.Underline = wdUnderlineSingle
      .Execute Replace:=wdReplaceAll
    End With
  Next A
  Exit Sub
  
FixUnderlinesError:
  Err.Source = strCleanup & "FixUnderlines"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Function zz_errorChecks()
  On Error GoTo zz_errorChecksError
    zz_errorChecks = False

    '-----test if backtick style tag already exists
    Set activeRng = activeDoc.Range

    Dim existingTagArray(3) As String   ' number of items in array should be declared here
    Dim B As Long
    Dim foundBad As Boolean
    foundBad = False

    existingTagArray(1) = "`[0-9]`"
    existingTagArray(2) = "`[A-Z]|"
    existingTagArray(3) = "|[A-Z]`"

    For B = 1 To UBound(existingTagArray())
        With activeRng.Find
            .ClearFormatting
            .Text = existingTagArray(B)
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With

        If activeRng.Find.Execute Then foundBad = True: Exit For
    Next

    If foundBad = True Then                'If activeRng.Find.Execute Then
'        MsgBox "Something went wrong! The macro cannot be run on Document:" & vbNewLine & "'" & activeDoc & _
'            "'" & vbNewLine & vbNewLine & "Please contact Digital Workflow group for support, I am sure they will " & _
'            "be happy to help.", , "Error Code: 3"
        zz_errorChecks = True
        Err.Raise MacError.err_BacktickCharFound
    End If
    Exit Function
    
zz_errorChecksError:
  Err.Source = strCleanup & "zz_errorChecks"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function


Private Sub NumRangeHyphens(StoriesInDoc As Variant)
  On Error GoTo NumRangeHyphensError
  ' convert hyphens in number ranges to en-dashes,
  ' but doesn't change hyphens in URLs or phone numbers

  ' tag URLs w/ macmillan style, so we can avoid later
  Call StyleAllHyperlinks(StoriesInUse:=StoriesInDoc)
  
  Dim strFindStart As String
  Dim strFindEnd As String
  Dim strTag As String
  Dim strFindWhat As String
  Dim strReplaceWith As String
  Dim strLinkStyle As String
  Dim activeRange As Range
  Dim kStory As Long
  
' Patterns to find and replace
' exclude start-with-hyphen or end-with-hyphen to exclude phone numbers, ISBN,
' SSN, and the like
  strFindStart = "([!\-]<[0-9]@)"
  strFindEnd = "([0-9]@>[!\-])"
  strTag = "`|url|`"
  
  strFindWhat = strFindStart & "\-" & strFindEnd
  strReplaceWith = "\1" & strTag & "\2"
  ' Macmillan URL style name
  strLinkStyle = "span hyperlink (url)"
  
'  For kStory = LBound(StoriesInDoc) To UBound(StoriesInDoc)
'    Set activeRange = activeDoc.StoryRanges(StoriesInDoc(kStory))
  Set activeRange = activeDoc.Range
  
  Call genUtils.GeneralHelpers.zz_clearFind
  With activeRange.Find
' Find each thing that is also a URL
' and replace hyphen with tags
    .Text = strFindWhat
    .Replacement.Text = strReplaceWith
    .Style = strLinkStyle
    .MatchWildcards = True
    .Wrap = wdFindStop
    .Forward = True
    .Format = True
    .Execute Replace:=wdReplaceAll

    ' Find the rest and replace with en-dash
    strReplaceWith = "\1^=\2"
    .ClearFormatting
    .Replacement.Text = strReplaceWith
    .Execute Replace:=wdReplaceAll
    
    ' Replace url tags w/ original hyphen
    strFindWhat = strFindStart & strTag & strFindEnd
    strReplaceWith = "\1-\2"
    
    .ClearFormatting
    .Replacement.ClearFormatting
    .Replacement.Text = strReplaceWith
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
  End With
  Exit Sub
  
NumRangeHyphensError:
  Err.Source = strCleanup & "NumRangeHyphens"
  If ErrorChecker(Err, strLinkStyle) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub


Private Sub CleanSomeSymbols(StoryTypes As WdStoryType)
' Remove formatting from some symbols
 On Error GoTo CleanSomeSymbolsError
  
  Dim activeRange As Range
  Set activeRange = activeDoc.StoryRanges(StoryTypes)
  
  Dim arrSymbols(1 To 3) As String
  Dim X As Long
  
  arrSymbols(1) = "^0174"    ' (r) registered trademark symbol
  arrSymbols(2) = "^0169"    ' (c) copyright symbol
  arrSymbols(3) = "^0153"    ' TM trademark symbol
  
  Call genUtils.GeneralHelpers.zz_clearFind
  ' Just removing superscript for right now
  For X = LBound(arrSymbols) To UBound(arrSymbols)
    With activeRange.Find
      .Text = arrSymbols(X)
      .Replacement.Text = "^&"
      .Format = True
      .Replacement.Font.Superscript = False
      .MatchWildcards = True
      .Execute Replace:=wdReplaceAll
    End With
  Next X
  Exit Sub
  
CleanSomeSymbolsError:
  Err.Source = strCleanup & "CleanSomeSymbols"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

' ===== ShapeDelete ===========================================================
' Deletes shape objects other than Text Boxes, Comments, and Tables. Should
' probably add more handling for text boxes before releasing to main cleanup
' macro.

Private Sub ShapeDelete()
  On Error GoTo ShapeDeleteError
  Dim objInlineShp As InlineShape
  Dim lngCount As Long
  Dim A As Long
  Dim objShape As Shape
  Dim typeShape As MsoShapeType

' Convert any InlineShapes to regular Shapes
  If activeDoc.InlineShapes.Count > 0 Then
    For Each objInlineShp In activeDoc.InlineShapes
      objInlineShp.ConvertToShape
    Next objInlineShp
  End If
  
' Note that TEXT BOXES are SHAPES!!
  lngCount = activeDoc.Shapes.Count
  If lngCount > 0 Then
    For A = lngCount To 1 Step -1
      Set objShape = activeDoc.Shapes(A)
      typeShape = objShape.Type
'      DebugPrint typeShape
      Select Case typeShape
        Case msoTextBox
          ' Do nothing for now, will need to find a solution eventually
        Case Else
          objShape.Delete
      End Select
    Next A
  End If
  Exit Sub
  
ShapeDeleteError:
  Err.Source = strCleanup & "ShapeDelete"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub
