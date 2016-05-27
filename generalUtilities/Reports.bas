Attribute VB_Name = "Reports"
' ====== PURPOSE =================
' Checks that manuscript styles follow Macmillan best practices
' If no Macmillan styles are used, just produces a list of styles in use

' ====== DEPENDENCIES ============
' 1. Manuscript must be styled with Macmillan custom styles to generate full report.
' 2. Requires ProgressBar userform
' 3. Requires SharedMacros be installed in same template.


Option Explicit
Option Base 1

Sub BookmakerReqs()
    Call MakeReport(torDOTcom:=True)
End Sub


Sub MacmillanStyleReport()
    Call MakeReport(torDOTcom:=False)
End Sub


Private Sub MakeReport(torDOTcom As Boolean)
    '-----------------------------------------------------------
    
    'Created by Erica Warren - erica.warren@macmillan.com
    
    '=================================================
    '''''              Timer Start                  '|
    'Dim StartTime As Double                         '|
    'Dim SecondsElapsed As Double                    '|
                                                    '|
    '''''Remember time when macro starts            '|
    'StartTime = Timer                               '|
    '=================================================
    
    '------------check for endnotes and footnotes--------------------------
    Dim arrStories() As Variant
    arrStories = StoryArray
    
    ' ======= Run startup checks ========
    ' True means a check failed (e.g., doc protection on)
    If StartupSettings(StoriesUsed:=arrStories) = True Then
        Call Cleanup
        Exit Sub
    End If
    
    '--------Progress Bar------------------------------
    'Percent complete and status for progress bar (PC) and status bar (Mac)
    'Requires ProgressBar custom UserForm and Class
    Dim sglPercentComplete As Single
    Dim strStatus As String
    Dim strTitle As String
    
    'First status shown will be randomly pulled from array, for funzies
    Dim funArray() As String
    ReDim funArray(1 To 10)      'Declare bounds of array here
    
    If torDOTcom = True Then
        funArray(1) = "* Is this thing on?..."
        funArray(2) = "* Are we there yet?..."
        funArray(3) = "* Zapping space invaders..."
        funArray(4) = "* Leaping over tall buildings in a single bound..."
        funArray(5) = "* Taking a quick nap..."
        funArray(6) = "* Taking the stairs..."
        funArray(7) = "* Partying like it's 1999..."
        funArray(8) = "* Waiting in line at Shake Shack..."
        funArray(9) = "* Revving engines..."
        funArray(10) = "* Thanks for running the Bookmaker Macro!"
    Else
        funArray(1) = "* Now is the winter of our discontent, made glorious summer by these Word Styles..."
        funArray(2) = "* What's in a name? Word Styles by any name would smell as sweet..."
        funArray(3) = "* A horse! A horse! My Word Styles for a horse!"
        funArray(4) = "* Be not afraid of Word Styles. Some are born with Styles, some achieve Styles, and some have Styles thrust upon 'em..."
        funArray(5) = "* All the world's a stage, and all the Word Styles merely players..."
        funArray(6) = "* To thine own Word Styles be true, and it must follow, as the night the day, thou canst not then be false to any man..."
        funArray(7) = "* To Style, or not to Style: that is the question..."
        funArray(8) = "* Word Styles, Word Styles! Wherefore art thou Word Styles?..."
        funArray(9) = "* Some Cupid kills with arrows, some with Word Styles..."
        funArray(10) = "* What light through yonder window breaks? It is the east, and Word Styles are the sun..."
    End If
    
    Dim X As Integer
    
    'Rnd returns random number between (0,1], rest of expression is to return an integer (1,10)
    Randomize           'Sets seed for Rnd below to value of system timer
    X = Int(UBound(funArray()) * Rnd()) + 1
    
    'Debug.Print x
    If torDOTcom = True Then
        strTitle = "Bookmaker Requirements Macro"
    Else
        strTitle = "Macmillan Style Report"
    End If
    
    sglPercentComplete = 0.02
    strStatus = funArray(X)
    
    Dim oProgressBkmkr As ProgressBar
    Set oProgressBkmkr = New ProgressBar    ' Triggers Initialize

    oProgressBkmkr.Title = strTitle
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    
    '-------remove "span ISBN (isbn)" style from letters, spaces, parens, etc.-------------------
    '-------because it should just be applied to the isbn numerals and hyphens-------------------
    Call ISBNcleanup
    
    '-------Count number of occurences of each required style----
    sglPercentComplete = 0.03
    strStatus = "* Counting required styles..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    Dim styleCount() As Variant
    
    styleCount = CountReqdStyles()
    
    If styleCount(1) = 100 Then     'Then count got stuck in a loop, gave message to user in last function
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' ------- Clear character formatting from Chapter Numbers ---------
    sglPercentComplete = 0.07
    strStatus = "* Cleaning up Chapter Numbers..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    If styleCount(4) > 0 Then
        Call ChapNumCleanUp
    End If
    
    
    '------------Convert unapproved headings to correct heading-------
    sglPercentComplete = 0.09
    strStatus = "* Correcting heading styles..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    ' If certain styles (oldStyle) appear by themselves, converts to
    ' the approved solo style (newStyle)
    
    If styleCount(4) > 0 And styleCount(5) = 0 Then
        Call FixSectionHeadings(oldStyle:="Chap Number (cn)", newStyle:="Chap Title (ct)")
    End If
    
    If styleCount(9) > 0 And styleCount(8) = 0 Then
        Call FixSectionHeadings(oldStyle:="Part Number (pn)", newStyle:="Part Title (pt)")
    End If
    
    If styleCount(11) > 0 And styleCount(10) = 0 Then
        Call FixSectionHeadings(oldStyle:="FM Title (fmt)", newStyle:="FM Head (fmh)")
    End If
    
    If styleCount(13) > 0 And styleCount(12) = 0 Then
        Call FixSectionHeadings(oldStyle:="BM Title (bmt)", newStyle:="BM Head (bmh)")
    End If
    

    
    '--------Get title/author/isbn/imprint text from document-----------
    sglPercentComplete = 0.11
    Application.ScreenUpdating = True
    strStatus = "* Getting book metadata from manuscript..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    Dim strMetadata As String
    strMetadata = GetMetadata
    
    '-------------------Get Illustrations List from Document-----------
    sglPercentComplete = 0.15
    strStatus = "* Getting list of illustrations..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    Dim strIllustrationsList As String
    strIllustrationsList = IllustrationsList
        
    '-------------------Get list of good and bad styles from document---------
    sglPercentComplete = 0.18
    strStatus = "* Getting list of styles in use..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    Dim arrGoodBadStyles() As Variant
    Dim strGoodStylesList As String
    Dim strBadStylesList As String
                
    'returns array with 2 elements, 1: good styles list, 2: bad styles list
    arrGoodBadStyles = GoodBadStyles(Tor:=torDOTcom, ProgressBar:=oProgressBkmkr, Status:=strStatus, ProgTitle:=strTitle, _
        Stories:=arrStories)
    strGoodStylesList = arrGoodBadStyles(1)
    'Debug.Print strGoodStylesList
    strBadStylesList = arrGoodBadStyles(2)
        
    'Error checking: if no good styles are in use, just return list of all styles in use, not other checks
    Dim blnTemplateUsed As Boolean
    Dim strSearchPattern As String
    ' Searching for "Footnote Text" or "Endnote Text" followed by page number, then
    ' followed by anything NOT including a close bracket. If there are other Mac styles
    ' it won't select the whole string
    
    strSearchPattern = "[EF]{1}[dnot]{4}[eot]{2,} Text -- p. [0-9]{1,}[!\)]{1,}"
    
    If strGoodStylesList = vbNullString Then
        blnTemplateUsed = False
    ' Test if good styles are just Endnote Text and Footnote Text
    ElseIf PatternMatch(SearchPattern:=strSearchPattern, SearchText:=strGoodStylesList, WholeString:=True) = True Then
        blnTemplateUsed = False
    Else
        blnTemplateUsed = True
    End If
    
    'If template not used, just returns list of styles in use
    If blnTemplateUsed = False Then
        strGoodStylesList = StylesInUse(ProgressBar:=oProgressBkmkr, Status:=strStatus, ProgTitle:=strTitle, Stories:=arrStories)
        strBadStylesList = ""
    End If
    
    '-------------------Create list of errors----------------------------
    sglPercentComplete = 0.98
    strStatus = "* Checking styles for errors..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    Dim strErrorList As String
    
    If blnTemplateUsed = True Then
        strErrorList = CreateErrorList(badStyles:=strBadStylesList, arrStyleCount:=styleCount, blnTor:=torDOTcom)
        'strErrorList = "testing"
    Else
        strErrorList = ""
    End If
    
    '------Create Report Text-------------------------------
    sglPercentComplete = 0.99
    strStatus = "* Creating report file..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    Dim strReportText As String
    strReportText = CreateReportText(blnTemplateUsed, strErrorList, strMetadata, strIllustrationsList, strGoodStylesList)
    
    
    ' Create Report File -------------------------------------
    Dim strSuffix As String
    If torDOTcom = True Then
        strSuffix = "BookmakerReport" ' suffix for the report file
    Else
        strSuffix = "StyleReport"
    End If
    
    Call CreateTextFile(strText:=strReportText, suffix:=strSuffix)
    
    '-------------Go back to original settings-----------------
    sglPercentComplete = 1
    strStatus = "* Finishing up..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    Call Cleanup
    
    Unload oProgressBkmkr
    
    '============================================================================
    '----------------------Timer End-------------------------------------------
    ''''Determine how many seconds code took to run
      'SecondsElapsed = Round(Timer - StartTime, 2)
    
    ''''Notify user in seconds
      'Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
    '============================================================================

End Sub

Private Function GoodBadStyles(Tor As Boolean, ProgressBar As ProgressBar, Status As String, ProgTitle As String, Stories() As Variant) As Variant
    'Creates a list of Macmillan styles in use
    'And a separate list of non-Macmillan styles in use
    
    Dim TheOS As String
    TheOS = System.OperatingSystem
    Dim sglPercentComplete As Single
    Dim strStatus As String
    
    Dim activeDoc As Document
    Set activeDoc = ActiveDocument
    Dim stylesGood() As String
    Dim stylesGoodLong As Long
    stylesGoodLong = 400                                    'could maybe reduce this number
    ReDim stylesGood(stylesGoodLong)
    Dim stylesBad() As String
    ReDim stylesBad(1 To 100) 'could maybe reduce this number too
    Dim styleGoodCount As Integer
    Dim styleBadCount As Integer
    Dim styleBadOverflow As Boolean
    Dim activeParaCount As Integer
    Dim J As Integer, K As Integer, L As Integer
    Dim paraStyle As String
    '''''''''''''''''''''
    Dim activeParaRange As Range
    Dim pageNumber As Integer
    Dim A As Long
    
    
    'Alter built-in Normal (Web) style temporarily (later, maybe forever?)
    ActiveDocument.Styles("Normal (Web)").NameLocal = "_"
    
    '----------Collect all styles being used-------------------------------
    styleGoodCount = 0
    styleBadCount = 0
    styleBadOverflow = False
    activeParaCount = activeDoc.Paragraphs.Count
    For J = 1 To activeParaCount
        
        'All Progress Bar statements for PC only because won't run modeless on Mac
        If J Mod 100 = 0 Then
        
            'Percent complete and status for progress bar (PC) and status bar (Mac)
            sglPercentComplete = (((J / activeParaCount) * 0.45) + 0.18)
            strStatus = "* Checking paragraph " & J & " of " & activeParaCount & " for Macmillan styles..." & _
                        vbCr & Status
            
            'Debug.Print sglPercentComplete
            Call UpdateBarAndWait(Bar:=ProgressBar, Status:=strStatus, Percent:=sglPercentComplete)
        End If
        
        For A = LBound(Stories()) To UBound(Stories())
            If J <= ActiveDocument.StoryRanges(Stories(A)).Paragraphs.Count Then
                paraStyle = activeDoc.StoryRanges(Stories(A)).Paragraphs(J).Style
                Set activeParaRange = activeDoc.StoryRanges(Stories(A)).Paragraphs(J).Range
                pageNumber = activeParaRange.Information(wdActiveEndPageNumber)                 'alt: (wdActiveEndAdjustedPageNumber)
                    
                'If InStrRev(paraStyle, ")", -1, vbTextCompare) Then        'ALT calculation to "Right", can speed test
                If Right(paraStyle, 1) = ")" Then
CheckGoodStyles:
                    For K = 1 To styleGoodCount
                        'Debug.Print Left(stylesGood(K), InStrRev(stylesGood(K), " --") - 1)
                        ' "Left" function because now stylesGood includes page number, so won't match paraStyle
                        If paraStyle = Left(stylesGood(K), InStrRev(stylesGood(K), " --") - 1) Then
                        K = styleGoodCount                              'stylereport bug fix #1    v. 3.1
                            Exit For                                        'stylereport bug fix #1   v. 3.1
                        End If                                              'stylereport bug fix #1   v. 3.1
                    Next K
                    
                    If K = styleGoodCount + 1 Then
                        styleGoodCount = K
                        ReDim Preserve stylesGood(1 To styleGoodCount)
                        stylesGood(styleGoodCount) = paraStyle & " -- p. " & pageNumber
                    End If
                
                Else
                    
                    If paraStyle = "Endnote Text" Or paraStyle = "Footnote Text" Then
                        GoTo CheckGoodStyles
                    Else
                        For L = 1 To styleBadCount
                            'If paraStyle = stylesBad(L) Then Exit For                  'Not needed, since we want EVERY instance of bad style
                        Next L
                        If L > 100 Then                                                 ' Exits if more than 100 bad paragraphs
                                styleBadOverflow = True
                                stylesBad(100) = "** WARNING: More than 100 paragraphs with bad styles found." & vbNewLine & vbNewLine
                            Exit For
                        End If
                        If L = styleBadCount + 1 Then
                            styleBadCount = L
            
                            stylesBad(styleBadCount) = "** ERROR: Non-Macmillan style on page " & pageNumber & _
                                " (Paragraph " & J & "):  " & paraStyle & vbNewLine & vbNewLine
                        End If
                     End If
                End If
            End If
        Next A
    Next J
    
    Status = "* Checking paragraphs for Macmillan styles..." & vbCr & Status
    
    'Change Normal (Web) back (if you want to)
    ActiveDocument.Styles("Normal (Web),_").NameLocal = "Normal (Web)"
    
    ' DON'T sort styles alphabetically, per request from PE
'    'Sort good styles
'    If K <> 0 Then
'    ReDim Preserve stylesGood(1 To styleGoodCount)
'    WordBasic.SortArray stylesGood()
'    End If
    
    'Create single string for good styles
    Dim strGoodStyles As String
    
    If styleGoodCount = 0 Then
        strGoodStyles = ""
    Else
        For K = LBound(stylesGood()) To UBound(stylesGood())
            strGoodStyles = strGoodStyles & stylesGood(K) & vbCrLf
        Next K
    End If
    
    'Debug.Print strGoodStyles
    
    If styleBadCount > 0 Then
        'Create single string for bad styles
        Dim strBadStyles As String
        ReDim Preserve stylesBad(1 To styleBadCount)
        For L = LBound(stylesBad()) To UBound(stylesBad())
            strBadStyles = strBadStyles & stylesBad(L)
        Next L
    Else
        strBadStyles = ""
    End If
    
    'Debug.Print strBadStyles
    
    '-------------------get list of good character styles--------------
    
    Dim charStyles As String
    Dim styleNameM(1 To 21) As String        'declare number in array
    Dim M As Integer
    
    styleNameM(1) = "span italic characters (ital)"
    styleNameM(2) = "span boldface characters (bf)"
    styleNameM(3) = "span small caps characters (sc)"
    styleNameM(4) = "span underscore characters (us)"
    styleNameM(5) = "span superscript characters (sup)"
    styleNameM(6) = "span subscript characters (sub)"
    styleNameM(7) = "span bold ital (bem)"
    styleNameM(8) = "span smcap ital (scital)"
    styleNameM(9) = "span smcap bold (scbold)"
    styleNameM(10) = "span symbols (sym)"
    styleNameM(11) = "span accent characters (acc)"
    styleNameM(12) = "span cross-reference (xref)"
    styleNameM(13) = "span hyperlink (url)"
    styleNameM(14) = "span material to come (tk)"
    styleNameM(15) = "span carry query (cq)"
    styleNameM(16) = "span preserve characters (pre)"
    styleNameM(17) = "span strikethrough characters (str)"
    styleNameM(18) = "bookmaker keep together (kt)"
    styleNameM(19) = "span ISBN (isbn)"
    styleNameM(20) = "span symbols ital (symi)"
    styleNameM(21) = "span symbols bold (symb)"
    
    
    
    For M = 1 To UBound(styleNameM())
        
        'Percent complete and status for progress bar (PC) and status bar (Mac)
        sglPercentComplete = (((M / UBound(styleNameM())) * 0.13) + 0.63)
        strStatus = "* Checking for " & styleNameM(M) & " styles..." & vbCr & Status
    
        Call UpdateBarAndWait(Bar:=ProgressBar, Status:=strStatus, Percent:=sglPercentComplete)
        
        On Error GoTo ErrHandler
        
        'Move selection back to start of document
        Selection.HomeKey Unit:=wdStory
        
        'Need to do Selection.Find for char styles. Range.Find won't work.
        'We only need to find a style once to add it to the list
        'Search through the main text story here
        With Selection.Find
            .Style = ActiveDocument.Styles(styleNameM(M))
            .Wrap = wdFindContinue
            .Format = True
            .Execute
        End With
        
        If Selection.Find.Found = True Then
            charStyles = charStyles & styleNameM(M) & vbNewLine
        'Else not present in main text story
        Else
            ' So check if there are footnotes
            If ActiveDocument.Footnotes.Count > 0 Then
                'If there are footnotes, select the footnote text
                ActiveDocument.StoryRanges(wdFootnotesStory).Select
                'Search the new selection for the style
                With Selection.Find
                    .Style = ActiveDocument.Styles(styleNameM(M))
                    .Wrap = wdFindContinue
                    .Format = True
                    .Execute
                End With
            
                If Selection.Find.Found = True Then
                    charStyles = charStyles & styleNameM(M) & vbNewLine
                ' Else didn't find style in footnotes, check endnotes
                Else
                    GoTo CheckEndnotes
                End If
            Else
CheckEndnotes:
                ' Check if there are endnotes in the document
                If ActiveDocument.Endnotes.Count > 0 Then
                    ' If there are endnotes, select them
                    ActiveDocument.StoryRanges(wdEndnotesStory).Select
                    'Search the new selection for the style
                    With Selection.Find
                         .Style = ActiveDocument.Styles(styleNameM(M))
                         .Wrap = wdFindContinue
                         .Format = True
                         .Execute
                     End With
                        
                    If Selection.Find.Found = True Then
                        charStyles = charStyles & styleNameM(M) & vbNewLine
                    End If
                End If
            End If
        End If
NextLoop:
    Next M
    
    'Debug.Print charStyles
    
    Status = "* Checking character styles..." & vbCr & Status
    
    'Add character styles to Good styles list
    strGoodStyles = strGoodStyles & charStyles
    
    'If this is for the Tor.com Bookmaker toolchain, test if only those styles used
    Dim strTorBadStyles As String
    If Tor = True Then
        strTorBadStyles = BadTorStyles(ProgressBar2:=ProgressBar, StatusBar:=Status, ProgressTitle:=ProgTitle, Stories:=Stories)
        strBadStyles = strBadStyles & strTorBadStyles
    End If
    
    'Debug.Print strGoodStyles
    'Debug.Print strBadStyles
    
    'If only good styles are Endnote Text and Footnote text, then the template is not being used
    
    
    'Add both good and bad styles lists to an array to pass back to original sub
    Dim arrFinalLists() As Variant
    ReDim arrFinalLists(1 To 2)
    
    arrFinalLists(1) = strGoodStyles
    arrFinalLists(2) = strBadStyles
    
    GoodBadStyles = arrFinalLists
    
    Exit Function
    
ErrHandler:
    'Debug.Print Err.Number & " : " & Err.Description
    If Err.Number = 5834 Or Err.Number = 5941 Then
        Resume NextLoop
    End If
    
End Function


Private Function CreateErrorList(badStyles As String, arrStyleCount() As Variant, blnTor As Boolean) As String
    Dim errorList As String
    
    errorList = ""
    
    '--------------For reference----------------------
    'arrStyleCount(1) = "Titlepage Book Title (tit)"
    'arrStyleCount(2) = "Titlepage Author Name (au)"
    'arrStyleCount(3) = "span ISBN (isbn)"
    'arrStyleCount(4) = "Chap Number (cn)"
    'arrStyleCount(5) = "Chap Title (ct)"
    'arrStyleCount(6) = "Chap Title Nonprinting (ctnp)"
    'arrStyleCount(7) = "Titlepage Logo (logo)"
    'arrStyleCount(8) = "Part Title (pt)"
    'arrStyleCount(9) = "Part Number (pn)"
    'arrStyleCount(10) = "FM Head (fmh)"
    'arrStyleCount(11) = "FM Title (fmt)"
    'arrStyleCount(12) = "BM Head (bmh)"
    'arrStyleCount(13) = "BM Title (bmt)"
    'arrStyleCount(14) = "Illustration holder (ill)"
    'arrStyleCount(15) = "Illustration source (is)"
    '------------------------------------------------
    
    '=====================Generate errors based on number of required elements found==================
    
    'If Book Title = 0
    If arrStyleCount(1) = 0 Then errorList = errorList & "** ERROR: No styled title detected." & _
        vbNewLine & vbNewLine
    
    'If Book Title > 1
    If arrStyleCount(1) > 1 Then errorList = errorList & "** ERROR: Too many title paragraphs detected." _
        & " Only 1 allowed." & vbNewLine & vbNewLine
    
    'Check if page break before Book Title
    If arrStyleCount(1) > 0 Then errorList = errorList & CheckPrevStyle(findStyle:="Titlepage Book Title (tit)", _
        prevStyle:="Page Break (pb)")
    
    
    'If Author Name = 0
    If arrStyleCount(2) = 0 Then errorList = errorList & "** ERROR: No styled author name detected." _
        & vbNewLine & vbNewLine
    
    'If ISBN = 0
    If arrStyleCount(3) = 0 Then
        errorList = errorList & "** ERROR: No styled ISBN detected." _
        & vbNewLine & vbNewLine
    Else
        If blnTor = True Then
            'check for correct book type following ISBN, in parens.
            errorList = errorList & BookTypeCheck
        End If
    End If
    
    'If CN > 0 and CT = 0 (already fixed in FixSectionHeadings sub)
    If arrStyleCount(4) > 0 And arrStyleCount(5) = 0 Then errorList = errorList & _
        "** WARNING: Chap Number (cn) cannot be the main heading for" & vbNewLine _
        & vbTab & "a chapter. Every chapter must include Chapter Title (ct)" & vbNewLine _
        & vbTab & "style. Chap Number (cn) paragraphs have been converted to the" & vbNewLine _
        & vbTab & "Chap Title (ct) style." & vbNewLine & vbNewLine
    
    'If PN > 0 and PT = 0 (already fixed in FixSectionHeadings sub)
    If arrStyleCount(9) > 0 And arrStyleCount(8) = 0 Then errorList = errorList & _
        "** WARNING: Part Number (pn) cannot be the main heading for" & vbNewLine _
        & vbTab & "a section. Every part must include Part Title (pt)" & vbNewLine _
        & vbTab & "style. Part Number (pn) paragraphs have been converted" & vbNewLine _
        & vbTab & "to the Part Title (pt) style." & vbNewLine & vbNewLine
    
    'If FMT > 0 and FMH = 0 (already fixed in FixSectionHeadings sub)
    If arrStyleCount(11) > 0 And arrStyleCount(10) = 0 Then errorList = errorList & _
        "** WARNING: FM Title (fmt) cannot be the main heading for" & vbNewLine _
        & vbTab & "a section. Every front matter section must include" & vbNewLine _
        & vbTab & "the FM Head (fmh) style. FM Title (fmt) paragraphs" & vbNewLine _
        & vbTab & "have been converted to the FM Head (fmh) style." & vbNewLine & vbNewLine
    
    'If BMT > 0 and BMH = 0 (already fixed in FixSectionHeadings sub)
    If arrStyleCount(13) > 0 And arrStyleCount(12) = 0 Then errorList = errorList & _
        "** WARNING: BM Title (bmt) cannot be the main heading for" & vbNewLine _
        & vbTab & "a section. Every back matter section must incldue" & vbNewLine _
        & vbTab & "the BM Head (bmh) style. BM Title (bmt) paragraphs" & vbNewLine _
        & vbTab & "have been converted to the BM Head (bmh) style." & vbNewLine & vbNewLine
            
    'If no chapter opening paragraphs (CN, CT, or CTNP)
    If arrStyleCount(4) = 0 And arrStyleCount(5) = 0 And arrStyleCount(6) = 0 Then errorList = errorList _
        & "** ERROR: No tagged chapter openers detected. If your book does" & vbNewLine _
        & vbTab & "not have chapter openers, use the Chap Title Nonprinting" & vbNewLine _
        & vbTab & "(ctnp) style at the start of each section." & vbNewLine & vbNewLine
    
    'If CN > CT and CT > 0 (i.e., Not a CT for every CN)
    If arrStyleCount(4) > arrStyleCount(5) And arrStyleCount(5) > 0 Then errorList = errorList & _
        "** ERROR: More Chap Number (cn) paragraphs than Chap Title (ct)" & vbNewLine _
        & vbTab & "paragraphs found. Each Chap Number (cn) paragraph MUST be" & vbNewLine _
        & vbTab & "followed by a Chap Title (ct) paragraph." & vbNewLine & vbNewLine
    
    'If Imprint line = 0
    If arrStyleCount(7) = 0 Then errorList = errorList & "** WARNING: No styled Titlepage Logo (logo) line detected. " _
        & "If you would like a logo included on your titlepage, please add this style." & vbNewLine & vbNewLine
    
    'If Imprint Lline > 1
    If arrStyleCount(7) > 1 Then errorList = errorList & "** ERROR: Too many Imprint Line paragraphs" _
        & " detected. Only 1 allowed." & vbNewLine & vbNewLine
    
    'If only CTs because converted by macro check for a page break before
    If (arrStyleCount(4) > 0 And arrStyleCount(5) = 0) Then errorList = errorList & _
        CheckPrevStyle(findStyle:="Chap Title (ct)", prevStyle:="Page Break (pb)")
    
    'If only PTs (either originally or converted by macro) check for a page break before
    If (arrStyleCount(9) > 0 And arrStyleCount(8) = 0) Or (arrStyleCount(9) = 0 And arrStyleCount(8) > 0) _
        Then errorList = errorList & CheckPrevStyle(findStyle:="Part Title (pt)", prevStyle:="Page Break (pb)")
    
    'If only FMHs (either originally or converted by macro) check for a page break before
    If (arrStyleCount(11) > 0 And arrStyleCount(10) = 0) Or (arrStyleCount(11) = 0 And arrStyleCount(10) > 0) _
        Then errorList = errorList & CheckPrevStyle(findStyle:="FM Head (fmh)", prevStyle:="Page Break (pb)")
    
    'If only BMHs (either originally or converted by macro) check for a page break before
    If (arrStyleCount(13) > 0 And arrStyleCount(12) = 0) Or (arrStyleCount(13) = 0 And arrStyleCount(12) > 0) _
        Then errorList = errorList & CheckPrevStyle(findStyle:="BM Head (bmh)", prevStyle:="Page Break (pb)")
    
    'If only CTP, check for a page break before
    If arrStyleCount(4) = 0 And arrStyleCount(5) = 0 And arrStyleCount(6) > 0 Then errorList = errorList _
        & CheckPrevStyle(findStyle:="Chap Title Nonprinting (ctnp)", prevStyle:="Page Break (pb)")
            
    'If CNs <= CTs, then check that those 3 styles are in order
    If arrStyleCount(4) <= arrStyleCount(5) And arrStyleCount(4) > 0 Then errorList = errorList & CheckPrev2Paras("Page Break (pb)", _
        "Chap Number (cn)", "Chap Title (ct)")
    
    'If Illustrations and sources exist, check that source comes after Ill and Cap
    If blnTor = True Then
        If arrStyleCount(14) > 0 And arrStyleCount(15) > 0 Then errorList = errorList & _
            CheckPrev2Paras("Illustration holder (ill)", "Caption (cap)", "Illustration Source (is)")
        If CheckFileName = True Then errorList = errorList & _
            "**ERROR: Bookmaker can only accept file names that use" & vbNewLine & _
            "letters, numbers, hyphens, or underscores. Punctuation," & vbNewLine & _
            "spaces, and other special characters are not allowed." & vbNewLine & vbNewLine
    End If
    
    'Check that only heading styles follow page breaks
    errorList = errorList & CheckAfterPB
    
    ' Check that all CTNP have some text
    If arrStyleCount(6) > 0 Then errorList = errorList & CheckNonprintingText
    
    'Add bad styles to error message
    errorList = errorList & badStyles
    
    If errorList <> "" Then
        errorList = errorList & vbNewLine & "If you have any questions about how to handle these errors, " & vbNewLine & _
            "please contact workflows@macmillan.com." & vbNewLine
    End If
    
    'Debug.Print errorList
    
    CreateErrorList = errorList

End Function


Function CheckPrevStyle(findStyle As String, prevStyle As String) As String
    Dim jString As String
    Dim jCount As Integer
    Dim pageNum As Integer
    Dim intCurrentPara As Integer
    
    Application.ScreenUpdating = False
    
        'check if styles exist, else exit sub
        On Error GoTo ErrHandler:
        Dim keyStyle As Word.Style
    
        Set keyStyle = ActiveDocument.Styles(findStyle)
        Set keyStyle = ActiveDocument.Styles(prevStyle)
    
    jCount = 0
    jString = ""
    
    'Move selection to start of document
    Selection.HomeKey Unit:=wdStory
    
    'select paragraph with that style
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = ActiveDocument.Styles(findStyle)
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    
    Do While Selection.Find.Execute = True And jCount < 200            'jCount so we don't get an infinite loop
        jCount = jCount + 1
        
        'Get number of current pagaraph, because we get an error if try to select before 1st para
        Dim rParagraphs As Range
        Dim CurPos As Long
         
        Selection.Range.Select  'select current ran
        CurPos = ActiveDocument.Bookmarks("\startOfSel").Start
        Set rParagraphs = ActiveDocument.Range(Start:=0, End:=CurPos)
        intCurrentPara = rParagraphs.Paragraphs.Count
        
        'Debug.Print intCurrentPara
        
        If intCurrentPara > 1 Then
            'select preceding paragraph
            Selection.Previous(Unit:=wdParagraph, Count:=1).Select
            pageNum = Selection.Information(wdActiveEndPageNumber)
        
                'Check if preceding paragraph style is correct
                If Selection.Style <> prevStyle Then
                    jString = jString & "** ERROR: Missing or incorrect " & prevStyle & " style before " _
                        & findStyle & " style on page " & pageNum & "." & vbNewLine & vbNewLine
                End If
            
                'If you're searching for a page break before, also check if manual page break is in paragraph
                If prevStyle = "Page Break (pb)" Then
                    If InStr(Selection.Text, Chr(12)) = 0 Then
                        jString = jString & "** ERROR: Missing manual page break on page " & pageNum & "." _
                            & vbNewLine & vbNewLine
                    End If
                End If
            
                'Debug.Print jString
        
            'move the selection back to original paragraph, so it won't be
            'selected again on next search
            Selection.Next(Unit:=wdParagraph, Count:=1).Select
        End If
        
    Loop
    
    'Debug.Print jString
    
    CheckPrevStyle = jString
    
    Exit Function
    
ErrHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then       'style doesn't exist in document
        Exit Function
    End If
End Function

Function CheckAfterPB() As String
    Dim arrSecStartStyles() As String
    ReDim arrSecStartStyles(1 To 44)
    Dim kString As String
    Dim kCount As Integer
    Dim pageNumK As Integer
    Dim nextStyle As String
    Dim N As Integer
    Dim nCount As Integer
    
    Application.ScreenUpdating = False
    
    ' These are all styles allowed to follow a page break
    arrSecStartStyles(1) = "Chap Title (ct)"
    arrSecStartStyles(2) = "Chap Number (cn)"
    arrSecStartStyles(3) = "Chap Title Nonprinting (ctnp)"
    arrSecStartStyles(4) = "Halftitle Book Title (htit)"
    arrSecStartStyles(5) = "Titlepage Book Title (tit)"
    arrSecStartStyles(6) = "Copyright Text single space (crtx)"
    arrSecStartStyles(7) = "Copyright Text double space (crtxd)"
    arrSecStartStyles(8) = "Dedication (ded)"
    arrSecStartStyles(9) = "Ad Card Main Head (acmh)"
    arrSecStartStyles(10) = "Ad Card List of Titles (acl)"
    arrSecStartStyles(11) = "Part Title (pt)"
    arrSecStartStyles(12) = "Part Number (pn)"
    arrSecStartStyles(13) = "Front Sales Title (fst)"
    arrSecStartStyles(14) = "Front Sales Quote (fsq)"
    arrSecStartStyles(15) = "Front Sales Quote NoIndent (fsq1)"
    arrSecStartStyles(16) = "Epigraph - non-verse (epi)"
    arrSecStartStyles(17) = "Epigraph - verse (epiv)"
    arrSecStartStyles(18) = "FM Head (fmh)"
    arrSecStartStyles(19) = "Illustration holder (ill)"
    arrSecStartStyles(20) = "Page Break (pb)"
    arrSecStartStyles(21) = "FM Epigraph - non-verse (fmepi)"
    arrSecStartStyles(22) = "FM Epigraph - verse (fmepiv)"
    arrSecStartStyles(23) = "FM Head ALT (afmh)"
    arrSecStartStyles(24) = "Part Epigraph - non-verse (pepi)"
    arrSecStartStyles(25) = "Part Epigraph - verse (pepiv)"
    arrSecStartStyles(26) = "Part Contents Main Head (pcmh)"
    arrSecStartStyles(27) = "Poem Title (vt)"
    arrSecStartStyles(28) = "Recipe Head (rh)"
    arrSecStartStyles(29) = "Sub-Recipe Head (srh)"
    arrSecStartStyles(30) = "BM Head (bmh)"
    arrSecStartStyles(31) = "BM Head ALT (abmh)"
    arrSecStartStyles(32) = "Appendix Head (aph)"
    arrSecStartStyles(33) = "About Author Text (atatx)"
    arrSecStartStyles(34) = "About Author Text No-Indent (atatx1)"
    arrSecStartStyles(35) = "About Author Text Head (atah)"
    arrSecStartStyles(36) = "Colophon Text (coltx)"
    arrSecStartStyles(37) = "Colophon Text No-Indent (coltx1)"
    arrSecStartStyles(38) = "BOB Ad Title (bobt)"
    arrSecStartStyles(39) = "Series Page Heading (sh)"
    arrSecStartStyles(40) = "span small caps characters (sc)"
    arrSecStartStyles(41) = "span italic characters (ital)"
    arrSecStartStyles(42) = "Design Note (dn)"
    arrSecStartStyles(43) = "Front Sales Quote Head (fsqh)"
    arrSecStartStyles(44) = "Section Break (sbr)"
    
    kCount = 0
    kString = ""
    
    'Move selection to start of document
    Selection.HomeKey Unit:=wdStory
    
    On Error GoTo ErrHandler1
    
    'select paragraph styled as Page Break with manual page break inserted
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = "^m^p"
            .Replacement.Text = "^m^p"
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = ActiveDocument.Styles("Page Break (pb)")
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    
    Do While Selection.Find.Execute = True And kCount < 200            'jCount so we don't get an infinite loop
        kCount = kCount + 1
        nCount = 0
        'select following paragraph
        Selection.Next(Unit:=wdParagraph, Count:=1).Select
        nextStyle = Selection.Style
        pageNumK = Selection.Information(wdActiveEndPageNumber)
            
           For N = LBound(arrSecStartStyles()) To UBound(arrSecStartStyles())
                'Check if preceding paragraph style is correct
                If nextStyle <> arrSecStartStyles(N) Then
                    nCount = nCount + 1
                Else
                    Exit For
                End If
            Next N
                
            If nCount = UBound(arrSecStartStyles()) Then
                kString = kString & "** ERROR: " & nextStyle & " style on page " & pageNumK _
                    & " cannot follow Page Break (pb) style." & vbNewLine & vbNewLine
            End If
                    
        'Debug.Print kString
     
Err2Resume:
        
        'move the selection back to original paragraph, so it won't be
        'selected again on next search
        Selection.Previous(Unit:=wdParagraph, Count:=1).Select
    Loop
    
    'Debug.Print kString
    
    CheckAfterPB = kString
    
    Exit Function

ErrHandler1:
    If Err.Number = 5941 Or Err.Number = 5834 Then       'Style doesn't exist in document
        Exit Function
    End If
    
ErrHandler2:
    If Err.Number = 5941 Or Err.Number = 5834 Then       ' Style doesn't exist in document
        Resume Err2Resume
    End If

End Function


Private Function FixTrackChanges() As Boolean
    Dim N As Long
    Dim oComments As Comments
    Set oComments = ActiveDocument.Comments
    
    Application.ScreenUpdating = False
    
    FixTrackChanges = True
    
    Application.DisplayAlerts = False
    
    'See if there are tracked changes or comments in document
    On Error Resume Next
    Selection.HomeKey Unit:=wdStory   'start search at beginning of doc
    WordBasic.NextChangeOrComment       'search for a tracked change or comment. error if none are found.
    
    'If there are changes, ask user if they want macro to accept changes or cancel
    If Err = 0 Then
        If MsgBox("Bookmaker doesn't like comments or tracked changes, but it appears that you have some in your document." _
            & vbCr & vbCr & "Click OK to ACCEPT ALL CHANGES and DELETE ALL COMMENTS right now and continue with the Bookmaker Requirements Check." _
            & vbCr & vbCr & "Click CANCEL to stop the Bookmaker Requirements Check and deal with the tracked changes and comments on your own.", _
            273, "Are those tracked changes I see?") = vbCancel Then           '273 = vbOkCancel(1) + vbCritical(16) + vbDefaultButton2(256)
                FixTrackChanges = False
                Exit Function
        Else 'User clicked OK, so accept all tracked changes and delete all comments
            ActiveDocument.AcceptAllRevisions
            For N = oComments.Count To 1 Step -1
                oComments(N).Delete
            Next N
            Set oComments = Nothing
        End If
    End If
    
    On Error GoTo 0
    Application.DisplayAlerts = True
    
End Function

Private Function BadTorStyles(ProgressBar2 As ProgressBar, StatusBar As String, ProgressTitle As String, Stories() As Variant) As String
    'Called from GoodBadStyles sub if torDOTcom parameter is set to True.
    
    Dim paraStyle As String
    Dim activeParaCount As Integer
    
    Dim strCsvFileName As String
    Dim strLogInfo() As Variant
    ReDim strLogInfo(1 To 3)
    Dim strFullPathToCsv As String
    Dim arrTorStyles() As Variant
    Dim strLogDir As String
    Dim strPathToLogFile As String
    
    Dim intBadCount As Integer
    Dim activeParaRange As Range
    Dim pageNumber As Integer
    
    Dim N As Integer
    Dim M As Integer
    Dim strBadStyles As String
    Dim A As Long
    
    Dim TheOS As String
    TheOS = System.OperatingSystem
    Dim sglPercentComplete As Single
    Dim strStatus As String
    
    Application.ScreenUpdating = False
    
    
    ' This is the file we want to download
    strCsvFileName = "Styles_Bookmaker.csv"
    
    ' We need the info about the log file for any download
    strLogInfo() = CreateLogFileInfo(FileName:=strCsvFileName)
    strLogDir = strLogInfo(2)
    strPathToLogFile = strLogInfo(3)
    strFullPathToCsv = strLogDir & Application.PathSeparator & strCsvFileName
    
    ' download the list of good Tor styles from Confluence
    Dim downloadStyles As GitBranch
    ' switch to develop for testing
    downloadStyles = master
    
    If DownloadFromConfluence(DownloadSource:=downloadStyles, _
                                FinalDir:=strLogDir, _
                                LogFile:=strPathToLogFile, _
                                FileName:=strCsvFileName) = False Then
        ' If it's False, DL failed. Is a previous version there?
        If IsItThere(strFullPathToCsv) = False Then
            ' Sorry can't DL right now, no previous file in directory
            MsgBox "Sorry, I can't download the Bookmaker style info right now."
            Exit Function
        Else
            ' Can't DL new file but old one exists, let's use that
            MsgBox "I can't download the Bookmaker style info right now, so I'll just use the old info I have on file."
        End If
    End If
    
    
    'List of styles approved for use in Bookmaker
    'Organized by approximate frequency in manuscripts (most freq at top)
    'returned array is dimensioned with 1 column, need to specify row and column (base 0)
    arrTorStyles = LoadCSVtoArray(Path:=strFullPathToCsv, RemoveHeaderRow:=True, RemoveHeaderCol:=False)
    
    activeParaCount = ActiveDocument.Paragraphs.Count
    
    For N = 1 To activeParaCount
        
 
        If N Mod 100 = 0 Then
            'Percent complete and status for progress bar (PC) and status bar (Mac)
            sglPercentComplete = (((N / activeParaCount) * 0.1) + 0.76)
            strStatus = "* Checking paragraph " & N & " of " & activeParaCount & " for approved Bookmaker styles..." & vbCr & StatusBar
    
            Call UpdateBarAndWait(Bar:=ProgressBar2, Status:=strStatus, Percent:=sglPercentComplete)
        End If
        
        For A = LBound(Stories()) To UBound(Stories())
            If N <= ActiveDocument.StoryRanges(Stories(A)).Paragraphs.Count Then
                paraStyle = ActiveDocument.StoryRanges(Stories(A)).Paragraphs(N).Style
                'Debug.Print paraStyle
                
                If Right(paraStyle, 1) = ")" Then
                    'Debug.Print "Current paragraph is: " & paraStyle
                    On Error GoTo ErrHandler
                    
                    intBadCount = -1        ' -1 because the array is base 0
                    
                    For M = LBound(arrTorStyles()) To UBound(arrTorStyles())
                        'Debug.Print arrTorStyles(M, 0)
                        
                        If paraStyle <> arrTorStyles(M, 0) Then
                            intBadCount = intBadCount + 1
                        Else
                            Exit For
                        End If
                    Next M
                    
                    'Debug.Print intBadCount
                    If intBadCount = UBound(arrTorStyles()) Then
                        Set activeParaRange = ActiveDocument.StoryRanges(A).Paragraphs(N).Range
                        pageNumber = activeParaRange.Information(wdActiveEndPageNumber)
                        strBadStyles = strBadStyles & "** ERROR: Non-Bookmaker style on page " & pageNumber _
                            & " (Paragraph " & N & "):  " & paraStyle & vbNewLine & vbNewLine
                            'Debug.Print strBadStyles
                    End If
                
                End If
            End If
        Next A
ErrResume:
    
    Next N
    
    StatusBar = "* Checking paragraphs for approved Bookmaker styles..." & vbCr & StatusBar
    
    'Debug.Print strBadStyles
    
    BadTorStyles = strBadStyles
    Exit Function

ErrHandler:
    Debug.Print Err.Number & " " & Err.Description & " | " & Err.HelpContext
    If Err.Number = 5941 Or Err.Number = 5834 Then       'style is not in document
        Resume ErrResume
    End If

End Function

Private Function CountReqdStyles() As Variant
    Dim arrStyleName(1 To 15) As String                      ' Declare number of items in array
    Dim intStyleCount() As Variant
    ReDim intStyleCount(1 To 15) As Variant                  ' Delcare items in array. Must be dynamic to pass back to Sub
    
    Dim A As Long
    Dim xCount As Integer
    
    Application.ScreenUpdating = False
    
    arrStyleName(1) = "Titlepage Book Title (tit)"
    arrStyleName(2) = "Titlepage Author Name (au)"
    arrStyleName(3) = "span ISBN (isbn)"
    arrStyleName(4) = "Chap Number (cn)"
    arrStyleName(5) = "Chap Title (ct)"
    arrStyleName(6) = "Chap Title Nonprinting (ctnp)"
    arrStyleName(7) = "Titlepage Logo (logo)"
    arrStyleName(8) = "Part Title (pt)"
    arrStyleName(9) = "Part Number (pn)"
    arrStyleName(10) = "FM Head (fmh)"
    arrStyleName(11) = "FM Title (fmt)"
    arrStyleName(12) = "BM Head (bmh)"
    arrStyleName(13) = "BM Title (bmt)"
    arrStyleName(14) = "Illustration holder (ill)"
    arrStyleName(15) = "Illustration Source (is)"
    
    For A = 1 To UBound(arrStyleName())
        On Error GoTo ErrHandler
        intStyleCount(A) = 0
        With ActiveDocument.Range.Find
            .ClearFormatting
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = ActiveDocument.Styles(arrStyleName(A))
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        Do While .Execute(Forward:=True) = True And intStyleCount(A) < 100   ' < 100 to prevent infinite loop, especially if content controls in title or author blocks
            intStyleCount(A) = intStyleCount(A) + 1
        Loop
        End With
ErrResume:
    Next
    
                
    '------------Exit Sub if exactly 100 Titles counted, suggests hidden content controls-----
    If intStyleCount(1) = 100 Then
        
        MsgBox "Something went wrong!" & vbCr & vbCr & "It looks like you might have content controls (form fields or drop downs) in your document, but Word for Mac doesn't play nicely with these." _
        & vbCr & vbCr & "Try running this macro on a PC or contact workflows@macmillan.com for assistance.", vbCritical, "OH NO!!"
        Exit Function
        
    End If
    
    'For A = 1 To UBound(arrStyleName())
    '    Debug.Print arrStyleName(A) & ": " & intStyleCount(A) & vbNewLine
    'Next A
    
    CountReqdStyles = intStyleCount()
    Exit Function

ErrHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then
        intStyleCount(A) = 0
        Resume ErrResume
    End If
        
End Function

Private Sub FixSectionHeadings(oldStyle As String, newStyle As String)

    Application.ScreenUpdating = False

    'check if styles exist, else exit sub
    On Error GoTo ErrHandler:
    Dim keyStyle As Word.Style

    Set keyStyle = ActiveDocument.Styles(oldStyle)
    Set keyStyle = ActiveDocument.Styles(newStyle)

    'Move selection to start of document
    Selection.HomeKey Unit:=wdStory

        'Find paras styles as CN and change to CT style
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles(oldStyle)
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles(newStyle)
        With Selection.Find
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
        End With
        Selection.Find.Execute Replace:=wdReplaceAll

Exit Sub
    
ErrHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then 'the requested member of the collection does not exist (i.e., style doesn't exist)
        Exit Sub
    End If
    
End Sub

Private Function GetMetadata() As String
    Dim styleNameB(3) As String         ' must declare number of items in array here
    Dim bString(3) As String            ' and here
    Dim B As Integer
    Dim strTitleData As String
    
    Application.ScreenUpdating = False
    
    styleNameB(1) = "Titlepage Book Title (tit)"
    styleNameB(2) = "Titlepage Author Name (au)"
    styleNameB(3) = "span ISBN (isbn)"
    
    For B = 1 To UBound(styleNameB())
        bString(B) = GetText(styleNameB(B))
        If bString(B) <> vbNullString Then
            bString(B) = "** " & styleNameB(B) & " **" & vbNewLine & _
                        bString(B) & vbNewLine
        End If
        
        strTitleData = strTitleData & bString(B)
        
    Next B
                
    'Debug.Print strTitleData
    
    GetMetadata = strTitleData

End Function

Private Function IllustrationsList() As String
    Dim cString(1000) As String             'Max number of illustrations. Could be lower than 1000.
    Dim cCount As Integer
    Dim pageNumberC As Integer
    Dim strFullList As String
    Dim N As Integer
    Dim strSearchStyle As String
    
    Application.ScreenUpdating = False
    
    strSearchStyle = "Illustration holder (ill)"
    cCount = 0
    
    'Move selection to start of document
    Selection.HomeKey Unit:=wdStory
        
        ' Check if search style exists in document
        On Error GoTo ErrHandler
        Dim keyStyle As Style
        
        Set keyStyle = ActiveDocument.Styles(strSearchStyle)
    
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = ActiveDocument.Styles(strSearchStyle)
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    
    Do While Selection.Find.Execute = True And cCount < 1000            'cCount < 1000 so we don't get an infinite loop
        cCount = cCount + 1
        pageNumberC = Selection.Information(wdActiveEndPageNumber)
        
        'If paragraph return exists in selection, don't select last character (the last paragraph return)
        If InStr(Selection.Text, Chr(13)) > 0 Then
            Selection.MoveEnd Unit:=wdCharacter, Count:=-1
        End If
        
        cString(cCount) = "Page " & pageNumberC & ": " & Selection.Text & vbNewLine
        
        'If the next character is a paragraph return, add that to the selection
        'Otherwise the next Find will just select the same text with the paragraph return
        Selection.MoveEndWhile Cset:=Chr(13), Count:=wdForward
        
    Loop
    
    'Move selection back to start of document
    Selection.HomeKey Unit:=wdStory
    
    If cCount > 1000 Then
        MsgBox "You have more than 1,000 illustrations tagged in your manuscript." & vbNewLine & _
        "Please contact workflows@macmillan.com to complete your illustration list."
    End If
    
    If cCount = 0 Then
        cCount = 1
        cString(1) = "no illustrations detected" & vbNewLine
    End If
    
    For N = 1 To cCount
        strFullList = strFullList & cString(N)
    Next N
    
    'Debug.Print strFullList
    
    IllustrationsList = strFullList
    
    Exit Function

ErrHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then
        IllustrationsList = ""
        Exit Function
    End If

End Function

Function CheckPrev2Paras(StyleA As String, StyleB As String, StyleC As String) As String
    Dim strErrors As String
    Dim intCount As Integer
    Dim pageNum As Integer
    Dim intCurrentPara As Integer
    Dim strStyle1 As String
    Dim strStyle2 As String
    Dim strStyle3 As String
    
    Application.ScreenUpdating = False
    
        'check if styles exist, else exit sub
        On Error GoTo ErrHandler:
        Dim keyStyle As Word.Style
    
        Set keyStyle = ActiveDocument.Styles(StyleA)
        Set keyStyle = ActiveDocument.Styles(StyleB)
        Set keyStyle = ActiveDocument.Styles(StyleC)
    
    
    strErrors = ""
    
    'Move selection to start of document
    Selection.HomeKey Unit:=wdStory
    
    'select paragraph with that style
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = ActiveDocument.Styles(StyleC)
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    
    intCount = 0
    
    Do While Selection.Find.Execute = True And intCount < 300            'jCount < 300 so we don't get an infinite loop
        intCount = intCount + 1
        
        'Get number of current pagaraph, because we get an error if try to select before 1st para
        
        intCurrentPara = ActiveDocument.Range(0, Selection.Paragraphs(1).Range.End).Paragraphs.Count
        
        'Debug.Print intCurrentPara
        
        'Also determine if selection is the LAST paragraph of the document, for later
        Dim SelectionIncludesFinalParagraphMark As Boolean
        If Selection.Type = wdSelectionNormal And Selection.End = ActiveDocument.Content.End Then
            SelectionIncludesFinalParagraphMark = True
        Else
            SelectionIncludesFinalParagraphMark = False
        End If
        
        'Debug.Print intCurrentPara
        
        If intCurrentPara > 1 Then      'NOT first paragraph of document
            'select preceding paragraph
            Selection.Previous(Unit:=wdParagraph, Count:=1).Select
            pageNum = Selection.Information(wdActiveEndPageNumber)
        
                'Check if preceding paragraph style is correct
                If Selection.Style <> StyleA Then
                
                    If Selection.Style = StyleB Then
                        'select preceding paragraph again, see if it's prevStyle
                        Selection.Previous(Unit:=wdParagraph, Count:=1).Select
                        pageNum = Selection.Information(wdActiveEndPageNumber)
                        
                            If Selection.Style <> StyleA Then
                                strErrors = strErrors & "** ERROR: " & StyleB & " followed by " & StyleC & "" _
                                    & " on" & vbNewLine & vbTab & "page " & pageNum & " must be preceded by " _
                                    & StyleA & "." & vbNewLine & vbNewLine
                            Else
                                'If you're searching for a page break before, also check if manual page break is in paragraph
                                If StyleA = "Page Break (pb)" Then
                                    If InStr(Selection.Text, Chr(12)) = 0 Then
                                        strErrors = strErrors & "** ERROR: Missing manual page break on page " & pageNum & "." _
                                            & vbNewLine & vbNewLine
                                    End If
                                End If
                            End If
                            
                        Selection.Next(Unit:=wdParagraph, Count:=1).Select
                    Else
                    
                        strErrors = strErrors & "** ERROR: " & StyleC & " on page " _
                            & pageNum & " must be used after an" & vbNewLine & vbTab & StyleA & "." _
                                & vbNewLine & vbNewLine
                            
                    End If
                Else
                    'Make sure initial selection wasn't last paragraph, or else we'll error when trying to select after it
                    If SelectionIncludesFinalParagraphMark = False Then
                        'select follow paragraph again, see if it's a Caption
                        Selection.Next(Unit:=wdParagraph, Count:=2).Select
                        pageNum = Selection.Information(wdActiveEndPageNumber)
                            
                            If Selection.Style = StyleB Then
                                strErrors = strErrors & "** ERROR: " & StyleC & " style on page " & pageNum & " must" _
                                    & " come after " & StyleB & " style." & vbNewLine & vbNewLine
                            End If
                        Selection.Previous(Unit:=wdParagraph, Count:=2).Select
                    End If
                    
                    'If you're searching for a page break before, also check if manual page break is in paragraph
                    If StyleA = "Page Break (pb)" Then
                        If InStr(Selection.Text, Chr(12)) = 0 Then
                            strErrors = strErrors & "** ERROR: Missing manual page break on page " & pageNum & "." _
                                & vbNewLine & vbNewLine
                        End If
                    End If
                End If
            
                'Debug.Print strErrors
        
            'move the selection back to original paragraph, so it won't be
            'selected again on next search
            Selection.Next(Unit:=wdParagraph, Count:=1).Select
        
        Else 'Selection is first paragraph of the document
            strErrors = strErrors & "** ERROR: " & StyleC & " cannot be first paragraph of document." & vbNewLine & vbNewLine
        End If
        
    Loop
    
    '------------------------Search for Illustration holder and check previous paragraph--------------
    'Move selection to start of document
    Selection.HomeKey Unit:=wdStory
    
    'select paragraph with that style
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = ActiveDocument.Styles(StyleA)
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    
    intCount = 0
    
    Do While Selection.Find.Execute = True And intCount < 1000            'jCount < 1000 so we don't get an infinite loop
        intCount = intCount + 1
        
        'Get number of current pagaraph, because we get an error if try to select before 1st para
        intCurrentPara = ActiveDocument.Range(0, Selection.Paragraphs(1).Range.End).Paragraphs.Count
    
        If intCurrentPara > 1 Then      'NOT first paragraph of document
            'select preceding paragraph
            Selection.Previous(Unit:=wdParagraph, Count:=1).Select
            pageNum = Selection.Information(wdActiveEndPageNumber)
        
                'Check if preceding paragraph style is a Caption, which is not allowed
                If Selection.Style = StyleB Then
                    strErrors = strErrors & "** ERROR: " & StyleB & " on page " & pageNum & " must come after " _
                                    & StyleA & "." & vbNewLine & vbNewLine
                End If
                
            Selection.Next(Unit:=wdParagraph, Count:=1).Select
        End If
    Loop
    
    'Debug.Print strErrors
    
    CheckPrev2Paras = strErrors
    
    'Move selection back to start of document
    Selection.HomeKey Unit:=wdStory
    Exit Function

ErrHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then       'Style doesn't exist in document
        Exit Function
    End If

End Function

Private Function CreateReportText(TemplateUsed As Boolean, errorList As String, metadata As String, illustrations As String, goodStyles As String) As String

    Application.ScreenUpdating = False
    
    Dim strReportText As String
        
    If TemplateUsed = False Then
        strReportText = strReportText & vbNewLine
        strReportText = strReportText & "------------------------STYLES IN USE--------------------------" & vbNewLine
        strReportText = strReportText & "It looks like you aren't using the Macmillan style template." & vbNewLine
        strReportText = strReportText & "That's OK, but if you would like more info about your document," & vbNewLine
        strReportText = strReportText & "just attach the Macmillan style template and apply the styles" & vbNewLine
        strReportText = strReportText & "throughout the document." & vbNewLine
        strReportText = strReportText & vbNewLine
        strReportText = strReportText & goodStyles
    Else
        If errorList = "" Then
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & "                 CONGRATULATIONS! YOU PASSED!" & vbNewLine
            strReportText = strReportText & " But you're not done yet. Please check the info listed below." & vbNewLine
            strReportText = strReportText & vbNewLine
        Else
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & "                             OOPS!" & vbNewLine
            strReportText = strReportText & "     Problems were found with the styles in your document." & vbNewLine
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & "---------------------------- ERRORS ---------------------------" & vbNewLine
            strReportText = strReportText & errorList
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & vbNewLine
        End If
            strReportText = strReportText & "--------------------------- METADATA --------------------------" & vbNewLine
            strReportText = strReportText & "If any of the information below is wrong, please fix the" & vbNewLine
            strReportText = strReportText & "associated styles in the manuscript." & vbNewLine
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & metadata
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & "----------------------- ILLUSTRATION LIST ---------------------" & vbNewLine
        
            If illustrations <> "no illustrations detected" & vbNewLine Then
                strReportText = strReportText & "Verify that this list of illustrations includes only the file" & vbNewLine
                strReportText = strReportText & "names of your illustrations." & vbNewLine
                strReportText = strReportText & vbNewLine
            End If
        
            strReportText = strReportText & illustrations
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & "-------------------- MACMILLAN STYLES IN USE ------------------" & vbNewLine
            strReportText = strReportText & goodStyles
    End If

    CreateReportText = strReportText
    
End Function

Private Function StylesInUse(ProgressBar As ProgressBar, Status As String, ProgTitle As String, Stories() As Variant) As String
    'Creates a list of all styles in use, not just Macmillan styles
    'No list of bad styles
    'For use when no Macmillan template is attached
    
    Dim TheOS As String
    TheOS = System.OperatingSystem
    Dim sglPercentComplete As Single
    Dim strStatus As String
    
    Dim activeDoc As Document
    Set activeDoc = ActiveDocument
    Dim stylesGood() As String
    Dim stylesGoodLong As Long
    stylesGoodLong = 400                                    'could maybe reduce this number
    ReDim stylesGood(stylesGoodLong)
    Dim styleGoodCount As Integer
    Dim activeParaCount As Integer
    Dim J As Integer, K As Integer, L As Integer
    Dim paraStyle As String
    '''''''''''''''''''''
    Dim activeParaRange As Range
    Dim pageNumber As Integer
    Dim A As Long
    
    '----------Collect all styles being used-------------------------------
    styleGoodCount = 0
    activeParaCount = activeDoc.Paragraphs.Count
    For J = 1 To activeParaCount
        
        'All Progress Bar statements for PC only because won't run modeless on Mac
        If J Mod 100 = 0 Then
        
            'Percent complete and status for progress bar (PC) and status bar (Mac)
            sglPercentComplete = (((J / activeParaCount) * 0.12) + 0.86)
            strStatus = "* Checking paragraph " & J & " of " & activeParaCount & " for Macmillan styles..." & vbCr & Status
    
            Call UpdateBarAndWait(Bar:=ProgressBar, Status:=strStatus, Percent:=sglPercentComplete)
            
        End If
        
        For A = LBound(Stories()) To UBound(Stories())
            If J <= ActiveDocument.StoryRanges(Stories(A)).Paragraphs.Count Then
                paraStyle = activeDoc.StoryRanges(Stories(A)).Paragraphs(J).Style
                Set activeParaRange = activeDoc.StoryRanges(Stories(A)).Paragraphs(J).Range
                pageNumber = activeParaRange.Information(wdActiveEndPageNumber)                 'alt: (wdActiveEndAdjustedPageNumber)
        
                For K = 1 To styleGoodCount
                    ' "Left" function because now stylesGood includes page number, so won't match paraStyle
                    If paraStyle = Left(stylesGood(K), InStrRev(stylesGood(K), " --") - 1) Then
                        K = styleGoodCount                              'stylereport bug fix #1    v. 3.1
                        Exit For                                        'stylereport bug fix #1   v. 3.1
                    End If                                              'stylereport bug fix #1   v. 3.1
                Next K
                If K = styleGoodCount + 1 Then
                    styleGoodCount = K
                    stylesGood(styleGoodCount) = paraStyle & " -- p. " & pageNumber
                End If
            End If
        Next A
    Next J
    
    'Sort good styles
    If K <> 0 Then
    ReDim Preserve stylesGood(1 To styleGoodCount)
    WordBasic.SortArray stylesGood()
    End If
    
    'Create single string for good styles
    Dim strGoodStyles As String
    For K = LBound(stylesGood()) To UBound(stylesGood())
        strGoodStyles = strGoodStyles & stylesGood(K) & vbNewLine
    Next K
    
    'Debug.Print strGoodStyles
    
    StylesInUse = strGoodStyles

End Function

Private Sub ISBNcleanup()
'removes "span ISBN (isbn)" style from all but the actual ISBN numerals
    
    'check if that style exists, if not then exit sub
    On Error GoTo ErrHandler:
        Dim keyStyle As Word.Style
        Set keyStyle = ActiveDocument.Styles("span ISBN (isbn)")
    On Error GoTo 0
    
    Dim strISBNtextArray()
    ReDim strISBNtextArray(1 To 3)
    
    strISBNtextArray(1) = "-[!0-9]"     'any hyphen followed by any non-digit character
    strISBNtextArray(2) = "[!0-9]-"     'any hyphen preceded by any non-digit character
    strISBNtextArray(3) = "[!-0-9]"     'any character other than a hyphen or digit
    
    ' re: above--need to search for hyphens first, because if you lead with what is now 3, you
    ' remove the style from any characters around hyphens, so if you search for a hyphen next to
    ' a character later, it won't return anything because the whole string needs to have the
    ' style applied for it to be found.
    
    Dim G As Long
    For G = LBound(strISBNtextArray()) To UBound(strISBNtextArray())
        
        'Move selection to start of document
        Selection.HomeKey Unit:=wdStory

        With Selection.Find
            .ClearFormatting
            .Text = strISBNtextArray(G)
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = "span ISBN (isbn)"                     'find this style
            .Replacement.Style = "Default Paragraph Font"   'replace with this style
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        
        Selection.Find.Execute Replace:=wdReplaceAll
    
    Next G
    
Exit Sub
    
ErrHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then       'Style doesn't exist in document
        Exit Sub
    End If

End Sub

Private Function BookTypeCheck()
    ' Validates the book types listed following the ISBN on the copyright page.
    Dim intCount As Integer
    Dim strErrors As String
    Dim strBookTypes(1 To 7) As String
    Dim A As Long
    Dim blnMissing As Boolean
    Dim strISBN As String
    
    strBookTypes(1) = "trade paperback"
    strBookTypes(2) = "hardcover"
    strBookTypes(3) = "e-book"
    strBookTypes(4) = "ebook"
    strBookTypes(5) = "print on demand"
    strBookTypes(6) = "print-on-demand"
    strBookTypes(7) = "mass market paperback"
    
    'Move selection back to start of document
    Selection.HomeKey Unit:=wdStory

    On Error GoTo ErrHandler
    
    intCount = 0
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles("span ISBN (isbn)")
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Forward:=True) = True And intCount < 100   ' < 100 to precent infinite loop
            intCount = intCount + 1
            strISBN = Selection.Text
            'Record current selection because we need to return to it later
            ActiveDocument.Bookmarks.Add Name:="ISBN", Range:=Selection.Range
            
            Selection.Collapse Direction:=wdCollapseEnd
            Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            
            blnMissing = True
                For A = 1 To UBound(strBookTypes())
                    If InStr(Selection.Text, "(" & strBookTypes(A) & ")") > 0 Then
                        blnMissing = False
                        Exit For
                    End If
                Next A
            
            If blnMissing = True Then
                strErrors = strErrors & "** ERROR: Correct book type required in parentheses after" & vbNewLine & _
                    "ISBN " & strISBN & " on copyright page." _
                    & vbNewLine & vbNewLine
            End If
            
            'Now we need to return the selection to where it was above, or else we can't loop through selection.find
            If ActiveDocument.Bookmarks.Exists("ISBN") = True Then
                Selection.GoTo what:=wdGoToBookmark, Name:="ISBN"
                ActiveDocument.Bookmarks("ISBN").Delete
            End If
            
        Loop
    
    End With
    
    'Debug.Print strErrors
    BookTypeCheck = strErrors
    
    On Error GoTo 0
    Exit Function

ErrHandler:
    Debug.Print Err.Number & ": " & Err.Description
    If Err.Number = 5941 Or Err.Number = 5834 Then      ' style doesn't exist in document
        Exit Function
    End If
        
End Function

Private Function CheckNonprintingText()
    ' Verify that all "Chapter Title Nonprinting (ctnp)" paragraphs have some body text
    Dim iCount As Long
    Dim strBodyText As String
    Dim strErrors As String
    Dim pageNum As Long
    Dim intCount As Long

    
    'Move selection back to start of document
    Selection.HomeKey Unit:=wdStory

    On Error GoTo ErrHandler
    
    intCount = 0
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles("Chap Title Nonprinting (ctnp)")
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Forward:=True) = True And intCount < 1000   ' < 1000 to precent infinite loop
            intCount = intCount + 1
            strBodyText = Selection.Text

            pageNum = Selection.Information(wdActiveEndPageNumber)
            
'            'Record current selection because we need to return to it later
'            ActiveDocument.Bookmarks.Add Name:="CTNP", Range:=Selection.Range
'
'            Selection.Collapse Direction:=wdCollapseEnd
'            Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            
            If strBodyText = Chr(13) Then
                strErrors = strErrors & _
                    "** ERROR: Chap Title Nonprinting paragraph on page " & pageNum & " requires body text." & _
                    vbNewLine & vbNewLine
            End If

            
'            'Now we need to return the selection to where it was above, or else we can't loop through selection.find
'            If ActiveDocument.Bookmarks.Exists("ISBN") = True Then
'                Selection.GoTo what:=wdGoToBookmark, Name:="ISBN"
'                ActiveDocument.Bookmarks("ISBN").Delete
'            End If
            
        Loop
    
    End With
    
    CheckNonprintingText = strErrors
    
    Exit Function
    
ErrHandler:
        'Debug.Print Err.Number & ": " & Err.Description
    If Err.Number = 5941 Or Err.Number = 5834 Then      ' style doesn't exist in document
        Exit Function
    End If
    
End Function

Private Sub ChapNumCleanUp()
    ' Removes character styles from Chapter Number paragraphs
    Dim iCount As Long
    Dim strText As String
    Dim intCount As Long

    'Move selection back to start of document
    Selection.HomeKey Unit:=wdStory

    On Error GoTo ErrHandler
    
    intCount = 0
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles("Chap Number (cn)")
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Forward:=True) = True And intCount < 1000   ' < 1000 to precent infinite loop
            intCount = intCount + 1
            #If Mac Then
                ' Mac 2011 doesn't support ClearCharacterFormattingAll method
                ' And ClearFormatting removes paragraph formatting as well
                Selection.ClearFormatting
                Selection.Style = "Chap Number (cn)"
            #Else
                Selection.ClearCharacterAllFormatting
            #End If
        Loop
    
    End With
    
    
    Exit Sub
    
ErrHandler:
        'Debug.Print Err.Number & ": " & Err.Description
    If Err.Number = 5941 Or Err.Number = 5834 Then      ' style doesn't exist in document
        Exit Sub
    End If
End Sub


Private Function CheckFileName() As Boolean
' Returns error message if file name contains special characters

    Dim strDocName As String
    Dim strCheckChar As String
    Dim strAllGoodChars As String
    Dim lngNameLength As Long
    Dim R As Long
    Dim strErrorString As String
    
    CheckFileName = False
    
    ' Only alphanumeric, underscore and hyphen allowed in Bkmkr names
    ' Will do vbTextCompare later for case insensitive search
    strAllGoodChars = "ABCDEFGHIJKLMNOPQRSTUVWZYX1234567890_-"
    
    ' Get file name w/o extension
    strDocName = ActiveDocument.Name
    strDocName = Left(strDocName, InStrRev(strDocName, ".") - 1)
    
    lngNameLength = Len(strDocName)
    
    ' Loop: pull each char in file name, check if it appears in good char
    ' list. If it doesn't appear, then it's bad! So return True
    ' Error is same whether there is 1 or 100 bad chars, so exit as soon as
    ' one is found.
    
    For R = 1 To lngNameLength
        strCheckChar = Mid(strDocName, R, 1)
        If InStr(1, strAllGoodChars, strCheckChar, vbTextCompare) = 0 Then
            CheckFileName = True
            Exit Function
        End If
    Next R

End Function
