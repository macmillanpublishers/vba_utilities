Attribute VB_Name = "CharacterStyles"
' Created by Erica Warren -- erica.warren@macmillan.com
' Split off from MacmillanCleanupMacro: https://github.com/macmillanpublishers/Word-template/blob/master/macmillan/CleanupMacro.bas

' ======== PURPOSE ============================
' Applies Macmillan character styles to direct-styled text in current
' document

' ======== DEPENDENCIES =======================
' 1. Requires ProgressBar userform module
' 2. Requires GeneralHelpers module

' Note: have already used all numerals and capital letters for tagging,
' starting with lowercase letters. through a.

Option Explicit
Option Base 1

Private Const strCharStyles As String = "genUtils.CharacterStyles."
Private activeRng As Range

Public Function MacmillanCharStyles() As genUtils.Dictionary
  On Error GoTo MacmillanCharStylesError
  Dim dictReturn As genUtils.Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False
  
  Dim CharacterProgress As ProgressBar
  Set CharacterProgress = New ProgressBar
  
  CharacterProgress.Title = "Macmillan Character Styles Macro"
  DebugPrint "Starting Character Styles macro"
  
  Call genUtils.CharacterStyles.ActualCharStyles(oProgressChar:= _
    CharacterProgress, StartPercent:=0, TotalPercent:=1)
  
  dictReturn.Item("pass") = True
  Set MacmillanCharStyles = dictReturn
  Exit Function

MacmillanCharStylesError:
  Err.Source = strCharStyles & "MacmillanCharStyles"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function

Sub ActualCharStyles(oProgressChar As ProgressBar, StartPercent As Single, _
  TotalPercent As Single)
' Have to pass the ProgressBar so this can be run from within another macro
' StartPercent is the percentage the progress bar is at when this sub starts
' TotalPercent is the total percent of the progress bar that this sub will cover

  On Error GoTo ActualCharStylesError
'------------------Time Start-----------------
'Dim StartTime As Double
'Dim SecondsElapsed As Double

'Remember time when macro starts
'StartTime = Timer

' ------------check for endnotes and footnotes---------------------------------
  Dim stStories() As Variant
  stStories = genUtils.GeneralHelpers.StoryArray
    
' ======= Run startup checks ========
' True means a check failed (e.g., doc protection on)
  If genUtils.GeneralHelpers.StartupSettings(StoriesUsed:=stStories) = True Then
        
    Call genUtils.GeneralHelpers.Cleanup
    Exit Sub
  End If
  
' --------Progress Bar---------------------------------------------------------
' Percent complete and status for progress bar (PC) and status bar (Mac)
  Dim sglPercentComplete As Single
  Dim strStatus As String
  Dim strTitle As String
  
  'First status shown will be randomly pulled from array, for funzies
  Dim funArray() As String
  ReDim funArray(1 To 10)      'Declare bounds of array here
  
  funArray(1) = "* Mixing metaphors..."
  funArray(2) = "* Arguing about the serial comma..."
  funArray(3) = "* Un-mixing metaphors..."
  funArray(4) = "* Avoiding the passive voice..."
  funArray(5) = "* Ending sentences in prepositions..."
  funArray(6) = "* Splitting infinitives..."
  funArray(7) = "* Ooh, what an interesting manuscript..."
  funArray(8) = "* Un-dangling modifiers..."
  funArray(9) = "* Jazzing up author bio..."
  funArray(10) = "* Filling in plot holes..."
  
  Dim X As Integer
  
' Rnd returns random number between (0,1], rest of expression is to return an
' integer (1,10)
  Randomize           'Sets seed for Rnd below to value of system timer
  X = Int(UBound(funArray()) * Rnd()) + 1

' first number is percent of THIS macro completed
  sglPercentComplete = (0.09 * TotalPercent) + StartPercent
  strStatus = funArray(X)
  
' Calls ProgressBar.Increment mathod and waits for it to complete
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
' -----------Delete hidden text ------------------------------------------------
  Dim S As Long
  
  ' Note, if you don't delete hidden text, this macro turns it into reg. text.
  For S = 1 To UBound(stStories())
  If genUtils.GeneralHelpers.HiddenTextSucks(StoryType:=(stStories(S))) = _
    True Then
    ' Notify user maybe?
    End If
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind

' -------------- Clear formatting from paragraph marks ------------------------
' can cause errors
  
  For S = 1 To UBound(stStories())
    Call genUtils.GeneralHelpers.ClearPilcrowFormat(StoryType:=(stStories(S)))
  Next S


' ===================== Replace Local Styles Start ============================

' -----------------------Tag space break styles--------------------------------
  Call genUtils.GeneralHelpers.zz_clearFind
  
  sglPercentComplete = (0.18 * TotalPercent) + StartPercent
  strStatus = "* Preserving styled whitespace..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  For S = 1 To UBound(stStories())
    Call PreserveWhiteSpaceinBrkStylesA(StoryType:=(stStories(S)))
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind
  
' ----------------------------Fix hyperlinks-----------------------------------
  sglPercentComplete = (0.28 * TotalPercent) + StartPercent
  strStatus = "* Applying styles to hyperlinks..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  Call genUtils.GeneralHelpers.StyleAllHyperlinks(StoriesInUse:=stStories)
  
  Call genUtils.GeneralHelpers.zz_clearFind

' --------------------------Remove unstyled space breaks-----------------------
  sglPercentComplete = (0.39 * TotalPercent) + StartPercent
  strStatus = "* Removing unstyled breaks..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  For S = 1 To UBound(stStories())
    Call RemoveBreaks(StoryType:=(stStories(S)))
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind
  
' --------------------------Tag existing character styles----------------------
  sglPercentComplete = (0.52 * TotalPercent) + StartPercent
  strStatus = "* Tagging character styles..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  For S = 1 To UBound(stStories())
    Call TagExistingCharStyles(StoryType:=(stStories(S)))
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind
  
' -------------------------Tag direct formatting-------------------------------
  sglPercentComplete = (0.65 * TotalPercent) + StartPercent
  strStatus = "* Tagging direct formatting..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  ' allBkmkrStyles is a jagged array (array of arrays) to hold in-use Bookmaker styles.
  ' i.e., one array for each story. Must be Variant.
  Dim allBkmkrStyles() As Variant
  For S = 1 To UBound(stStories())
  'tag local styling, reset local styling, remove text highlights
    Call LocalStyleTag(StoryType:=(stStories(S)))
      
    ReDim Preserve allBkmkrStyles(1 To S)
    allBkmkrStyles(S) = TagBkmkrCharStyles(StoryType:=stStories(S))
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind

' ----------------------------Apply Macmillan character styles to tagged text--
  sglPercentComplete = (0.81 * TotalPercent) + StartPercent
  strStatus = "* Applying Macmillan character styles..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  For S = 1 To UBound(stStories())
    Call LocalStyleReplace(StoryType:=(stStories(S)), _
      BkmkrStyles:=allBkmkrStyles(S))
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind
  
' ---------------------------Remove tags from styled space breaks--------------
  sglPercentComplete = (0.95 * TotalPercent) + StartPercent
  strStatus = "* Cleaning up styled whitespace..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
  For S = 1 To UBound(stStories())
    Call PreserveWhiteSpaceinBrkStylesB(StoryType:=(stStories(S)))
  Next S
  
  Call genUtils.GeneralHelpers.zz_clearFind
    
' -------------- Tag un-styled paragraphs as TX / TX1 / COTX1 -----------------
' NOTE: must be done AFTER character styles, because if whole para has direct
' format it will be removed when apply style (but style won't be removed).
' This is total progress bar that will be covered in TagUnstyledText
  Dim sglTotalForText As Single
  sglTotalForText = TotalPercent - sglPercentComplete

  Call TagUnstyledText(objTagProgress:=oProgressChar, StartingPercent:= _
    sglPercentComplete, TotalPercent:=sglTotalForText, Status:=strStatus)

' Only tagging through main text story, because Endnotes story and Footnotes
' story should already be tagged at Endnote Text and Footnote Text by dafault
' when created

' ---------------------------Return settings to original-----------------------
  sglPercentComplete = TotalPercent + StartPercent
  strStatus = "* Finishing up..." & vbCr & strStatus
  
  Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=oProgressChar, _
    Status:=strStatus, Percent:=sglPercentComplete)
  
' If this is the whole macro, close out; otherwise calling macro will close it all down
  If TotalPercent = 1 Then
    Call genUtils.GeneralHelpers.Cleanup
    Unload oProgressChar
'        MsgBox "Macmillan character styles have been applied throughout your manuscript."
  End If
  

' ----------------------Timer End-------------------------------------------
' Determine how many seconds code took to run
' SecondsElapsed = Round(Timer - StartTime, 2)
  
' Notify user in seconds
'  DebugPrint "This code ran successfully in " & SecondsElapsed & " seconds"
  Exit Sub

ActualCharStylesError:
  Err.Source = strCharStyles & "ActualCharStyles"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub


Sub PreserveWhiteSpaceinBrkStylesA(StoryType As WdStoryType)
 On Error GoTo PreserveWhiteSpaceinBrkStylesAError:
  Set activeRng = activeDoc.StoryRanges(StoryType)

' Find/Replace (which we'll use later on) will not replace a paragraph mark
' in the first or last paragraph, so add dummy paragraphs here (with tags)
' that we can remove later on.

  activeRng.InsertBefore "``0``" & vbNewLine
  activeRng.InsertAfter vbNewLine & "``0``"
  
  Dim tagArray(13) As String
  Dim StylePreserveArray(13) As String
  Dim E As Long
  
  StylePreserveArray(1) = "Space Break (#)"
  StylePreserveArray(2) = "Space Break with Ornament (orn)"
  StylePreserveArray(3) = "Space Break with ALT Ornament (orn2)"
  StylePreserveArray(4) = "Section Break (sbr)"
  StylePreserveArray(5) = "Part Start (pts)"
  StylePreserveArray(6) = "Part End (pte)"
  StylePreserveArray(7) = "Page Break (pb)"
  StylePreserveArray(8) = "Space Break - 1-Line (ls1)"
  StylePreserveArray(9) = "Space Break - 2-Line (ls2)"
  StylePreserveArray(10) = "Space Break - 3-Line (ls3)"
  StylePreserveArray(11) = "Column Break (cbr)"
  StylePreserveArray(12) = "Design Note (dn)"
  StylePreserveArray(13) = "Bookmaker Page Break (br)"

' Only tag left side: we're searching for paragraph styles below, so each result
' will always end in ^13, which becomes our closing tag later. If we add a
' closing tag of our own, it gets added to the *following* paragraph, which
' complicates some later things we need to cleanup.

  tagArray(1) = "`1`^&"
  tagArray(2) = "`2`^&"
  tagArray(3) = "`3`^&"
  tagArray(4) = "`4`^&"
  tagArray(5) = "`5`^&"
  tagArray(6) = "`6`^&"
  tagArray(7) = "`7`^&"
  tagArray(8) = "`8`^&"
  tagArray(9) = "`9`^&"
  tagArray(10) = "`0`^&"
  tagArray(11) = "`L`^&"
  tagArray(12) = "`R`^&"
  tagArray(13) = "`N`^&"
  
  Call genUtils.GeneralHelpers.zz_clearFind
  For E = 1 To UBound(StylePreserveArray())
    With activeRng.Find
      .Text = "^13"
      .Replacement.Text = tagArray(E)
      .Wrap = wdFindContinue
      .Format = True
      .Style = StylePreserveArray(E)
      .MatchWildcards = True
      .Execute Replace:=wdReplaceAll
    End With
NextLoop:
  Next
  Exit Sub
    
PreserveWhiteSpaceinBrkStylesAError:
  ' skips tagging that style if it's missing from doc; if missing, obv nothing has that style
  'DebugPrint StylePreserveArray(e)
  '5834 "Item with specified name does not exist" i.e. style not present in doc
  '5941 item not available in collection
  If Err.Number = 5834 Or Err.Number = 5941 Then
      Resume NextLoop:
  End If
    
  Err.Source = strCharStyles & "PreserveWhiteSpaceinBrkStylesA"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Private Sub RemoveBreaks(StoryType As WdStoryType)
  On Error GoTo RemoveBreaksError
  Set activeRng = activeDoc.StoryRanges(StoryType)
    
  Dim wsFindArray(1 To 2) As String
  Dim wsReplaceArray(1 To 2) As String
  Dim Q As Long
  
' Remove page break and section break characters. Will have to re-evaluate
' before we move this to user-macro to determine what breaks to preserve
' (though section-start styles may make this moot).
  wsFindArray(1) = "^m"
  wsReplaceArray(1) = vbNullString

' Now that we've cleaned up errant page breaks, remove any blank paragraphs
  wsFindArray(2) = "^13{2,}"               '2 or more paragraphs
  wsReplaceArray(2) = "^p"


  Call genUtils.GeneralHelpers.zz_clearFind
  For Q = 1 To UBound(wsFindArray())
    With activeRng.Find
      .Text = wsFindArray(Q)
      .Replacement.Text = wsReplaceArray(Q)
      .Wrap = wdFindContinue
      .MatchWildcards = True
      .Execute Replace:=wdReplaceAll
    End With
  Next
  Exit Sub

RemoveBreaksError:
  Err.Source = strCharStyles & "RemoveBreaks"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Sub PreserveWhiteSpaceinBrkStylesB(StoryType As WdStoryType)
  On Error GoTo PreserveWhiteSpaceinBrkStylesBError
  
  Set activeRng = activeDoc.StoryRanges(StoryType)

' Now we need to remove our first/last dummy paragraphs
  Dim myRange(1 To 2) As Range
  Dim strText As String
  Dim strTag As String
  Dim strEnd As String
  Dim A As Long
  Set myRange(1) = activeDoc.Paragraphs.First.Range
  Set myRange(2) = activeDoc.Paragraphs.Last.Range
  
  For A = LBound(myRange) To UBound(myRange)
    strText = myRange(A).Text
    ' Validate that it is in fact the paragraph we added
    ' we added 5 chars, + new line char
    If Len(strText) >= 6 Then
    ' Separate our tag text from rest of para
      strTag = Left(strText, 5)
      strEnd = Right(strText, Len(strText) - 5)
    ' To be our added para, needs our tag AND new line char and nothing else
      If strTag = "``0``" And GeneralHelpers.IsNewLine(strEnd) = True Then
        myRange(A).Delete
      End If
    End If
  Next A

  Dim tagArrayB(13) As String
  Dim F As Long
    
  tagArrayB(1) = "`1`(^13)"
  tagArrayB(2) = "`2`(^13)"
  tagArrayB(3) = "`3`(^13)"
  tagArrayB(4) = "`4`(^13)"
  tagArrayB(5) = "`5`(^13)"
  tagArrayB(6) = "`6`(^13)"
  tagArrayB(7) = "`7`(^13)"
  tagArrayB(8) = "`8`(^13)"
  tagArrayB(9) = "`9`(^13)"
  tagArrayB(10) = "`0`(^13)"
  tagArrayB(11) = "`L`(^13)"
  tagArrayB(12) = "`R`(^13)"
  tagArrayB(13) = "`N`(^13)"

  Call genUtils.GeneralHelpers.zz_clearFind
  For F = 1 To UBound(tagArrayB())
    With activeRng.Find
      .Text = tagArrayB(F)
      .Replacement.Text = "\1"
      .Wrap = wdFindContinue
      .MatchWildcards = True
      .Execute Replace:=wdReplaceAll
    End With
  Next

' We also want to remove last para if it only contains a blank para, of any style
' Loop until we find a paragraph with text.
  Dim rngLast As Range
  Dim lngCount As Long
  
  lngCount = 0
  Do
    ' counter to prevent runaway loops
    lngCount = lngCount + 1
    Set rngLast = activeDoc.Paragraphs.Last.Range
    If GeneralHelpers.IsNewLine(rngLast.Text) = True Then
      rngLast.Delete
    Else
      Exit Do
    End If
  Loop Until lngCount = 20
  Exit Sub
  
PreserveWhiteSpaceinBrkStylesBError:
  Err.Source = strCharStyles & "PreserveWhiteSpaceinBrkStylesB"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Private Sub TagExistingCharStyles(StoryType As WdStoryType)
  On Error GoTo TagExistingCharStylesError
  Set activeRng = activeDoc.StoryRanges(StoryType)                        'this whole sub (except last stanza) is basically a v. 3.1 patch.  correspondingly updated sub name, call in main, and replacements go along with bold and common replacements
    
  Dim tagCharStylesArray(12) As String                                   ' number of items in array should be declared here
  Dim CharStylePreserveArray(12) As String              ' number of items in array should be declared here
  Dim D As Long
    
  CharStylePreserveArray(1) = "span hyperlink (url)"
  CharStylePreserveArray(2) = "span symbols (sym)"
  CharStylePreserveArray(3) = "span accent characters (acc)"
  CharStylePreserveArray(4) = "span cross-reference (xref)"
  CharStylePreserveArray(5) = "span material to come (tk)"
  CharStylePreserveArray(6) = "span carry query (cq)"
  CharStylePreserveArray(7) = "span key phrase (kp)"
  CharStylePreserveArray(8) = "span preserve characters (pre)"  'added v. 3.2

  CharStylePreserveArray(9) = "span ISBN (isbn)"  'added v. 3.7
  CharStylePreserveArray(10) = "span symbols ital (symi)"     'added v. 3.8
  CharStylePreserveArray(11) = "span symbols bold (symb)"
  CharStylePreserveArray(12) = "span run-in computer type (comp)"
  
  
  tagCharStylesArray(1) = "`H|^&|H`"
  tagCharStylesArray(2) = "`Z|^&|Z`"
  tagCharStylesArray(3) = "`Y|^&|Y`"
  tagCharStylesArray(4) = "`X|^&|X`"
  tagCharStylesArray(5) = "`W|^&|W`"
  tagCharStylesArray(6) = "`V|^&|V`"
  tagCharStylesArray(7) = "`T|^&|T`"
  tagCharStylesArray(8) = "`F|^&|F`"
  tagCharStylesArray(9) = "`Q|^&|Q`"
  tagCharStylesArray(10) = "`E|^&|E`"
  tagCharStylesArray(11) = "`G|^&|G`"
  tagCharStylesArray(12) = "`J|^&|J`"
  
  Call genUtils.GeneralHelpers.zz_clearFind
  For D = 1 To UBound(CharStylePreserveArray())
    With activeRng.Find
      .Replacement.Text = tagCharStylesArray(D)
      .Wrap = wdFindContinue
      .Format = True
      .Style = CharStylePreserveArray(D)
      .MatchWildcards = True
      .Execute Replace:=wdReplaceAll
    End With
NextLoop:
  Next
  Exit Sub
    
TagExistingCharStylesError:
' skips tagging that style if it's missing from doc;
' if missing, obv nothing has that style

' 5834 "Item with specified name does not exist" i.e. style not present in doc
' 5941 item is not present in collection
  If Err.Number = 5834 Or Err.Number = 5941 Then
    Resume NextLoop
  End If

  Err.Source = strCharStyles & "TagExistingCharStyles"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Private Sub LocalStyleTag(StoryType As WdStoryType)
  On Error GoTo LocalStyleTagError
  
  Set activeRng = activeDoc.StoryRanges(StoryType)
    
' ------------tag key styles---------------------------------------------------
  Dim tagStyleFindArray(11) As Boolean
  Dim tagStyleReplaceArray(11) As String
  Dim G As Long
  
  tagStyleFindArray(1) = False        'Bold
  tagStyleFindArray(2) = False        'Italic
  tagStyleFindArray(3) = False        'Underline
  tagStyleFindArray(4) = False        'Smallcaps
  tagStyleFindArray(5) = False        'Subscript
  tagStyleFindArray(6) = False        'Superscript
  tagStyleFindArray(7) = False        'Highlights
  ' note 8 - 10 are below
  tagStyleFindArray(11) = False       'Strikethrough
  
  tagStyleReplaceArray(1) = "`B|^&|B`"
  tagStyleReplaceArray(2) = "`I|^&|I`"
  tagStyleReplaceArray(3) = "`U|^&|U`"
  tagStyleReplaceArray(4) = "`M|^&|M`"
  tagStyleReplaceArray(5) = "`S|^&|S`"
  tagStyleReplaceArray(6) = "`P|^&|P`"
  tagStyleReplaceArray(8) = "`A|^&|A`"
  tagStyleReplaceArray(9) = "`C|^&|C`"
  tagStyleReplaceArray(10) = "`D|^&|D`"
  tagStyleReplaceArray(11) = "`a|^&|a`"

  Call genUtils.GeneralHelpers.zz_clearFind
  For G = 1 To UBound(tagStyleFindArray())
  
    tagStyleFindArray(G) = True
    
    If tagStyleFindArray(8) = True Then
      tagStyleFindArray(1) = True
      tagStyleFindArray(2) = True     'bold and italic
    End If
    
    If tagStyleFindArray(9) = True Then
      tagStyleFindArray(1) = True
      tagStyleFindArray(4) = True
      tagStyleFindArray(2) = False  'bold and smallcaps
    End If
    
    If tagStyleFindArray(10) = True Then
      tagStyleFindArray(2) = True
      tagStyleFindArray(4) = True
      tagStyleFindArray(1) = False 'smallcaps and italic
    End If
    
    If tagStyleFindArray(11) = True Then
      tagStyleFindArray(2) = False
      tagStyleFindArray(4) = False ' reset tags for strikethrough
    End If
    With activeRng.Find
      .Replacement.Text = tagStyleReplaceArray(G)
      .Wrap = wdFindContinue
      .Format = True
      .Font.Bold = tagStyleFindArray(1)
      .Font.Italic = tagStyleFindArray(2)
      .Font.Underline = tagStyleFindArray(3)
      .Font.SmallCaps = tagStyleFindArray(4)
      .Font.Subscript = tagStyleFindArray(5)
      .Font.Superscript = tagStyleFindArray(6)
      .Highlight = tagStyleFindArray(7)
      .Font.StrikeThrough = tagStyleFindArray(11)
      .Replacement.Highlight = False
      .MatchWildcards = True
      .Execute Replace:=wdReplaceAll
    End With
  
    tagStyleFindArray(G) = False

  Next
  Exit Sub

LocalStyleTagError:
  Err.Source = strCharStyles & "LocalStyleTag"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

Private Sub LocalStyleReplace(StoryType As WdStoryType, BkmkrStyles As Variant)
  On Error GoTo LocalStyleReplaceError
  
  Set activeRng = activeDoc.StoryRanges(StoryType)
  
  ' Determine if we need to do the bookmaker styles thing
  ' BkmkrStyles is an array of bookmaker character styles in use. If it's empty,
  ' there are none in use so we don't have to check
  
  Dim blnCheckBkmkr As Boolean
  
  If IsArrayEmpty(BkmkrStyles) = False Then
      blnCheckBkmkr = True
  Else
      blnCheckBkmkr = False
  End If
  
  '-------------apply styles to tags
  'number of items in array should = styles in LocalStyleTag + styles in TagExistingCharStyles
  Dim tagFindArray(22) As String              ' number of items in array should be declared here
  Dim tagReplaceArray(22) As String         'and here
  Dim H As Long
  
  tagFindArray(1) = "`B|(*)|B`"
  tagFindArray(2) = "`I|(*)|I`"
  tagFindArray(3) = "`U|(*)|U`"
  tagFindArray(4) = "`M|(*)|M`"
  tagFindArray(5) = "`H|(*)|H`"
  tagFindArray(6) = "`S|(*)|S`"
  tagFindArray(7) = "`P|(*)|P`"
  tagFindArray(8) = "`Z|(*)|Z`"
  tagFindArray(9) = "`Y|(*)|Y`"
  tagFindArray(10) = "`X|(*)|X`"
  tagFindArray(11) = "`W|(*)|W`"
  tagFindArray(12) = "`V|(*)|V`"
  tagFindArray(13) = "`T|(*)|T`"
  tagFindArray(14) = "`A|(*)|A`"                'v. 3.1 patch
  tagFindArray(15) = "`C|(*)|C`"                 'v. 3.1 patch
  tagFindArray(16) = "`D|(*)|D`"                       'v. 3.1 patch
  tagFindArray(17) = "`F|(*)|F`"
  tagFindArray(18) = "`Q|(*)|Q`"          'v. 3.7 added
  tagFindArray(19) = "`E|(*)|E`"
  tagFindArray(20) = "`G|(*)|G`"          'v. 3.8 added
  tagFindArray(21) = "`J|(*)|J`"
  tagFindArray(22) = "`a|(*)|a`"

  tagReplaceArray(1) = "span boldface characters (bf)"
  tagReplaceArray(2) = "span italic characters (ital)"
  tagReplaceArray(3) = "span underscore characters (us)"
  tagReplaceArray(4) = "span small caps characters (sc)"
  tagReplaceArray(5) = "span hyperlink (url)"
  tagReplaceArray(6) = "span subscript characters (sub)"
  tagReplaceArray(7) = "span superscript characters (sup)"
  tagReplaceArray(8) = "span symbols (sym)"
  tagReplaceArray(9) = "span accent characters (acc)"
  tagReplaceArray(10) = "span cross-reference (xref)"
  tagReplaceArray(11) = "span material to come (tk)"
  tagReplaceArray(12) = "span carry query (cq)"
  tagReplaceArray(13) = "span key phrase (kp)"
  tagReplaceArray(14) = "span bold ital (bem)"
  tagReplaceArray(15) = "span smcap bold (scbold)"
  tagReplaceArray(16) = "span smcap ital (scital)"
  tagReplaceArray(17) = "span preserve characters (pre)"
  tagReplaceArray(18) = "span ISBN (isbn)"                        'v. 3.7 added
  tagReplaceArray(19) = "span symbols ital (symi)"                ' v. 3.8 added
  tagReplaceArray(20) = "span symbols bold (symb)"                ' v. 3.8 added
  tagReplaceArray(21) = "span run-in computer type (comp)"
  tagReplaceArray(22) = "span strikethrough characters (str)"
  
  For H = LBound(tagFindArray()) To UBound(tagFindArray())
  
  ' ----------- bookmaker char styles ----------------------
    ' tag bookmaker line-ending character styles and
    ' adjust name if have additional styles applied
    ' because if you append "tighten" or "loosen" to
    ' regular style name, Bookmaker does that.
    If blnCheckBkmkr = True Then
      Dim Q As Long
      Dim qCount As Long
      Dim strAction As String
      Dim strNewName As String
      Dim strTag As String
      
      ' deal with bookmaker styles
      For Q = LBound(BkmkrStyles) To UBound(BkmkrStyles)
        ' replace bookmaker-tagged text with bookmaker styles
        strTag = "bk" & Format(Q, "0000")
        With activeRng.Find
          .ClearFormatting
          .Replacement.ClearFormatting
          .Text = "`" & strTag & "|(*)|" & strTag & "`"
          .Replacement.Text = "\1"
          .Wrap = wdFindContinue
          .Format = True
          .Replacement.Style = BkmkrStyles(Q)
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = True
          .MatchSoundsLike = False
          .MatchAllWordForms = False
          .Execute Replace:=wdReplaceAll
        End With

        'Move selection to start of document
        Selection.HomeKey Unit:=wdStory
    

        qCount = 0
        With Selection.Find
          .ClearFormatting
          .Replacement.ClearFormatting
          .Text = tagFindArray(H)
          .Replacement.Text = "\1"
          .Wrap = wdFindStop
          .Forward = True
          .Style = BkmkrStyles(Q)
          .Replacement.Style = tagReplaceArray(H)
          .Format = True
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = True
          .MatchSoundsLike = False
          .MatchAllWordForms = False

          Do While .Execute = True And qCount < 200
            qCount = qCount + 1
            .Execute Replace:=wdReplaceOne
            ' pull just action to add to style name
            ' always starts w/ "bookmaker ", but we want to include the space,
            ' hence start at 10
            strAction = Mid(BkmkrStyles(Q), 10, InStr(BkmkrStyles(Q), "(") - 11)
            strNewName = tagReplaceArray(H) & strAction
            
            ' Note these hybrid styles aren't in std template, so if they
            ' haven't been created in this doc yet, will error.
            Selection.Style = strNewName
          Loop
        End With
      Next Q
    End If
      
On Error GoTo ErrorHandler
    ' tag the rest of the character styles
    With activeRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = tagFindArray(H)
        .Replacement.Text = "\1"
        .Wrap = wdFindContinue
        .Format = True
        .Replacement.Style = tagReplaceArray(H)
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
  
NextLoop:
  Next
  Exit Sub

' TO DO: Move this to a more universal error handler for missing styles
' in main ErrorChecker function

ErrorHandler:
  
  Dim myStyle As Style
  
  If Err.Number = 5834 Or Err.Number = 5941 Then
    Select Case tagReplaceArray(H)
      
      'If style from LocalStyleTag is not present, add style
      Case "span boldface characters (bf)":
        Set myStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
          Type:=wdStyleTypeCharacter)
        With myStyle.Font
          .Shading.BackgroundPatternColor = wdColorLightTurquoise
          .Bold = True
        End With
        Resume
      
      Case "span italic characters (ital)"
        Set myStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
          Type:=wdStyleTypeCharacter)
        With myStyle.Font
          .Shading.BackgroundPatternColor = wdColorLightTurquoise
          .Italic = True
        End With
        Resume
          
      Case "span underscore characters (us)"
        Set myStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
          Type:=wdStyleTypeCharacter)
        With myStyle.Font
          .Shading.BackgroundPatternColor = wdColorLightTurquoise
          .Underline = wdUnderlineSingle
        End With
        Resume
      
      Case "span small caps characters (sc)"
        Set myStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
          Type:=wdStyleTypeCharacter)
        With myStyle.Font
          .Shading.BackgroundPatternColor = wdColorLightTurquoise
          .SmallCaps = False
          .AllCaps = True
          .Size = 9
        End With
        Resume
      
      Case "span subscript characters (sub)"
        Set myStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
          Type:=wdStyleTypeCharacter)
        With myStyle.Font
          .Shading.BackgroundPatternColor = wdColorLightTurquoise
          .Subscript = True
        End With
        Resume
          
      Case "span superscript characters (sup)"
        Set myStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
          Type:=wdStyleTypeCharacter)
        With myStyle.Font
          .Shading.BackgroundPatternColor = wdColorLightTurquoise
          .Superscript = True
        End With
        Resume

      Case "span bold ital (bem)"
        Set myStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
          Type:=wdStyleTypeCharacter)
        With myStyle.Font
          .Shading.BackgroundPatternColor = wdColorLightTurquoise
          .Bold = True
          .Italic = True
        End With
        Resume
          
      Case "span smcap bold (scbold)"
        Set myStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
          Type:=wdStyleTypeCharacter)
        With myStyle.Font
          .Shading.BackgroundPatternColor = wdColorLightTurquoise
          .SmallCaps = False
          .AllCaps = True
          .Size = 9
          .Bold = True
        End With
        Resume

      Case "span smcap ital (scital)"
        Set myStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), _
          Type:=wdStyleTypeCharacter)
        With myStyle.Font
          .Shading.BackgroundPatternColor = wdColorLightTurquoise
          .SmallCaps = False
          .AllCaps = True
          .Size = 9
          .Italic = True
        End With
        Resume
                
      Case "span strikethrough characters (str)"
        Set myStyle = activeDoc.Styles.Add(Name:=tagReplaceArray(H), Type:=wdStyleTypeCharacter)
        With myStyle.Font
          .Shading.BackgroundPatternColor = wdColorLightTurquoise
          .StrikeThrough = True
        End With
        Resume
            
    ' Else just skip if not from direct formatting
      Case Else
        Resume NextLoop:
        
    End Select
  End If
    
  Exit Sub

LocalStyleReplaceError:
'    DebugPrint Err.Number & ": " & Err.Description
'    DebugPrint "New name: " & strNewName
'    DebugPrint "Old name: " & tagReplaceArray(h)
  
  Dim myStyle2 As Style
  
  If Err.Number = 5834 Or Err.Number = 5941 Then

    Set myStyle2 = activeDoc.Styles.Add(Name:=strNewName, _
        Type:=wdStyleTypeCharacter)
        
On Error GoTo ErrorHandler
    ' If the original style did not exist yet, will error here
    ' but ErrorHandler will add the style
    myStyle2.BaseStyle = tagReplaceArray(H)
    ' Then go back to BkmkrError so further errors will route
    ' correctly
On Error GoTo LocalStyleReplaceError
    Resume
  Else
    Err.Source = strCharStyles & "LocalStyleReplace"
    If ErrorChecker(Err) = False Then
      Resume
    Else
      Call genUtils.GeneralHelpers.GlobalCleanup
    End If
  End If
    
End Sub


Private Function TagBkmkrCharStyles(StoryType As Variant) As Variant
  On Error GoTo TagBkmkrCharStylesError
'    Set activeRng = activeDoc.Range
  Set activeRng = activeDoc.StoryRanges(StoryType)
    
' Will need to loop through stories as well
' And be a function that returns an array

  Dim objStyle As Style
  Dim strBkmkrNames() As String
  Dim Z As Long
    
' Loop through all styles to get array of bkmkr styles in document
' NOTE! The .InUse property does NOT mean "in use in the document"; it means
' "any custom style or any modified built-in style". Ugh. Anyway, now we
' have to loop through all styles to see if bookmaker styles are present,
' then search for each of those styles to see if they are in use.
    
  For Each objStyle In activeDoc.Styles
    ' If char style with "bookmaker" in name is in use...
    ' binary compare is default, but adding here to be clear that we are doing
    ' a CASE SENSITIVE search, because "Bookmaker" is only for Paragraph styles,
    ' which we don't want to mess with.
    If InStr(1, objStyle.NameLocal, "bookmaker", vbBinaryCompare) <> 0 And _
      objStyle.Type = wdStyleTypeCharacter Then
'      DebugPrint StoryType & ": " & objStyle.NameLocal
      Selection.HomeKey Unit:=wdStory
      ' Now see if it's being used ...
      With Selection.Find
        .ClearFormatting
        .Text = ""
        .Style = objStyle.NameLocal
        .Wrap = wdFindContinue
        .Format = True
        .Forward = True
        .Execute
      End With
      
      If Selection.Find.Found = True Then
        '... add it to an array
        Z = Z + 1
        ReDim Preserve strBkmkrNames(1 To Z)
        strBkmkrNames(Z) = objStyle.NameLocal
      End If
    End If
  Next objStyle

  If IsArrayEmpty(strBkmkrNames) = True Then
    TagBkmkrCharStyles = strBkmkrNames
    Exit Function
  End If

' Tag in-use bkmkr styles
' Make sure if text also has formatting,
' the tags do not have it...
Dim X As Long
  Dim strTag As String
  Dim strAction As String
  Dim lngCount As Long

  For X = LBound(strBkmkrNames) To UBound(strBkmkrNames)
    strTag = "bk" & Format(X, "0000")
'        DebugPrint strTag
    
    With activeRng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Replacement.Text = "`" & strTag & "|^&|" & strTag & "`"
      .Wrap = wdFindContinue
      .Format = True
      .Style = strBkmkrNames(X)
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = True
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
  Next

  '-------------Reset everything -- clears all direct formatting!
  activeRng.Font.Reset
  
  ' return array of in-use bookmaker styles so we can tag later
  TagBkmkrCharStyles = strBkmkrNames
  Exit Function
  
TagBkmkrCharStylesError:
  Err.Source = strCharStyles & "TagBkmkrCharStyles"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Function

Private Sub TagUnstyledText(objTagProgress As ProgressBar, StartingPercent _
  As Single, TotalPercent As Single, Status As String)
  
  On Error GoTo TagUnstyledTextError
' Make sure we're always working with the right document
  Dim thisDoc As Document
  Set thisDoc = activeDoc

  ' Rename built-in style that has parens
  thisDoc.Styles("Normal (Web)").NameLocal = "_"

  Dim lngParaCount As Long
  Dim A As Long
  Dim strCurrentStyle As String
  Dim strTX As String
  Dim strTX1 As String
  Dim strNewStyle As String
  Dim strParaStatus As String
  Dim sglStartingPercent As Single
  Dim sglTotalPercent As Single
  Dim strNextStyle As String
  Dim strNextNextStyle As String
  Dim strCOTX1 As String
  Dim sglPercentComplete As Single

' Making these variables so we don't get any input errors with the style names t/o
  strTX = "Text - Standard (tx)"
  strTX1 = "Text - Std No-Indent (tx1)"
  strCOTX1 = "Chap Opening Text No-Indent (cotx1)"

  lngParaCount = thisDoc.Paragraphs.Count

  Dim myStyle As Style ' For error handlers

' Loop through all paras, tag any w/o close parens as TX or TX1
' (or COTX1 if following chap opener)
  For A = 1 To lngParaCount

    If A Mod 100 = 0 Then
      ' Increment progress bar
      sglPercentComplete = (((A / lngParaCount) * TotalPercent) + StartingPercent)
      strParaStatus = "* Tagging non-Macmillan paragraphs with Text - " & _
          "Standard (tx): " & A & " of " & lngParaCount & vbNewLine & Status
      Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=objTagProgress, Status:=strParaStatus, _
          Percent:=sglPercentComplete)
    End If

    strCurrentStyle = thisDoc.Paragraphs(A).Range.ParagraphStyle
'    DebugPrint strCurrentStyle

  ' tag all non-Macmillan-style paragraphs with standard Macmillan styles
  ' Macmillan styles all end in close parens
    If Right(strCurrentStyle, 1) <> ")" Then    ' it's not a Macmillan style
    ' If flush left, make No-Indent
      If thisDoc.Paragraphs(A).FirstLineIndent = 0 Then
          strNewStyle = strTX1
      Else
          strNewStyle = strTX
      End If
  
    ' Change the style of the paragraph in question
    ' This is where it will error if no style present
      thisDoc.Paragraphs(A).Style = strNewStyle
    
    ElseIf A < lngParaCount Then ' it is already a Macmillan style
    ' but can't check next para if it's the last para
    
    ' is it a chap head?
      If InStr(strCurrentStyle, "(cn)") > 0 Or _
        InStr(strCurrentStyle, "(ct)") > 0 Or _
        InStr(strCurrentStyle, "(ctnp)") > 0 Then

        strNextStyle = thisDoc.Paragraphs(A + 1).Range.ParagraphStyle

      ' is the next para non-Macmillan (and thus should be COTX1)
        If Right(strNextStyle, 1) <> ")" Then     ' it's not a Macmillan style
        ' so it should be COTX1
        ' Will error if this style not present in doc
          strNewStyle = strCOTX1
          thisDoc.Paragraphs(A + 1).Style = strNewStyle
        Else ' it IS a Macmillan style too
        ' it IT a chap opener? (can have CN followed by CT)
          If InStr(strNextStyle, "(cn)") > 0 Or _
            InStr(strNextStyle, "(ct)") > 0 Or _
            InStr(strNextStyle, "(ctnp)") > 0 Then

            strNextNextStyle = thisDoc.Paragraphs(A + 2).Range.ParagraphStyle

            If Right(strNextNextStyle, 1) <> ")" Then ' it's not Macmillan
            ' so it should be COTX1
              strNewStyle = strCOTX1
              thisDoc.Paragraphs(A + 2).Style = strNewStyle
            End If
          End If
        End If
      End If
    End If
  Next A

  ' Change Normal (Web) back
  thisDoc.Styles("Normal (Web),_").NameLocal = "Normal (Web)"

  Exit Sub

TagUnstyledTextError:
  Err.Source = strCharStyles & "TagUnstyledText"
  If ErrorChecker(Err, strNewStyle) = False Then
    Resume
  Else
    Call genUtils.GeneralHelpers.GlobalCleanup
  End If
End Sub

