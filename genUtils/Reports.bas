Attribute VB_Name = "Reports"
' ====== PURPOSE =================
' Variety of style/formatting checks for Macmillan styles

' ====== DEPENDENCIES ============
' 1. Manuscript must be styled with Macmillan custom styles to generate full
' report.
' 2. Requires genUtils be referenced from calling project.


' =============================================================================
'     DECLARATIONS
' =============================================================================

Option Explicit
Option Base 1

Private Const strReports As String = "genUtils.Reports."

' store info from `book_info.json` file
Private dictBookInfo As genUtils.Dictionary
' store initial paragraph loop info
Private dictStyles As genUtils.Dictionary
' store acceptable heading styles
Private dictHeadings As genUtils.Dictionary
' store style-to-section conversion
Private dictSections As genUtils.Dictionary
' also path to write alerts to
Private strAlertFile As String

Private Enum BookInfo
  bk_Title = 1
  bk_Authors = 2
  bk_ISBN = 3
End Enum


' ===== ReportsStartup ========================================================
' Set some global vars, check some things. Probably should be a class Initialize
' eventually.

Public Function ReportsStartup(DocPath As String, AlertPath As String) _
  As genUtils.Dictionary
  On Error GoTo ReportsStartupError
' Get this first, in case we have an error early:
  strAlertFile = AlertPath

' The .ps1 that calls this macro also opens the file, so should already be
' part of the Documents collection, but we'll check anyway.
    If genUtils.GeneralHelpers.IsOpen(DocPath) = False Then
    Documents.Open (DocPath)
  End If

' Set object variable (global scope, so other procedures can use)
  Set activeDoc = Documents(DocPath)
  
' Check for `book_info.json` file, read into global dictionary variable
  Dim strInfoPath As String
  strInfoPath = activeDoc.Path & Application.PathSeparator & "book_info.json"
  If GeneralHelpers.IsItThere(strInfoPath) = True Then
    Set dictBookInfo = genUtils.ClassHelpers.ReadJson(strInfoPath)
  Else
    Err.Raise MacError.err_FileNotThere
  End If

' Report success of initialize
  Dim dictReturn As genUtils.Dictionary
  Set dictReturn = New genUtils.Dictionary
  
  If dictBookInfo Is Nothing Then
    dictReturn.Add "pass", False
  Else
    dictReturn.Add "pass", True
  End If
  
  Set ReportsStartup = dictReturn
  
  Exit Function
ReportsStartupError:
  Err.Source = strReports & "ReportsStartup"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.ReportsTerminate
  End If
End Function


' ===== ReportsTerminate ======================================================
' Things to do if we have to terminate the macro early due to an error. To be
' called if ErrorChecker returns false. Again, some day would work better as
' a class with a legit Class_Terminate procedure.

Private Sub ReportsTerminate()
  Dim lngErrNumber As Long
  Dim strErrDescription As String
  Dim strErrSource As String
  
  lngErrNumber = Err.Number
  strErrDescription = Err.Description
  strErrSource = Err.Source

' Get current Err values before new `On Error` (which clear Err object)
  On Error GoTo ReportsTerminateError

' Write error file to active doc dir
  If strAlertFile = vbNullString Then
    If Not activeDoc Is Nothing Then
      strAlertFile = activeDoc.Path
    Else
      strAlertFile = ActiveDocument.Path
    End If
    strAlertFile = strAlertFile & Application.PathSeparator & "ALERT_" & _
        Format(Now, "yyyy-mm-dd_hh:mm") & ".txt"
  End If
  
' Create error message from Err values
  Dim strAlert As String
  strAlert = "=========================================" & vbNewLine & _
    Now & " | " & strErrSource & vbNewLine & _
    lngErrNumber & ": " & strErrDescription
  
  
' if we can write a file there, write alert message
  If genUtils.GeneralHelpers.ParentDirExists(strAlertFile) = True Then
    Dim FileNum As Long
    FileNum = FreeFile()
    Open strAlertFile For Append As #FileNum
    Print #FileNum, strAlert
    Close #FileNum
  Else
    ' just in case it stays...
    Debug.Print strAlert
  End If
  
' Kill global objects
  If Not dictBookInfo Is Nothing Then
    Set dictBookInfo = Nothing
  End If
  If Not dictStyles Is Nothing Then
    Set dictStyles = Nothing
  End If
  If Not dictHeadings Is Nothing Then
    Set dictHeadings = Nothing
  End If
  If Not dictSections Is Nothing Then
    Set dictSections = Nothing
  End If
  If Not activeDoc Is Nothing Then
    Set activeDoc = Nothing
  End If

  ' Close all open documents
  Dim objDoc As Document
  Dim strExt As String
  For Each objDoc In Documents
    ' don't close any macro templates, might be running code.
    strExt = VBA.Right(objDoc.Name, InStr(StrReverse(objDoc.Name), "."))
    If strExt <> ".dotm" Then
      objDoc.Close saveValue
    End If
  Next objDoc

' Do NOT use `Exit Sub` before, since we ALWAYS want this to terminate.
ReportsTerminateError:
  End
End Sub


' ===== StyleCheck ============================================================
' Call this from origin project. Performs variety of style checks, returns
' dictionary containing results of various tests or whatever. Use private
' global variable to store the `dictStyles` object to access by later
' procedures.

Public Function StyleCheck(Optional FixUnstyled As Boolean = True) As _
  genUtils.Dictionary

  On Error GoTo StyleDictionaryError
  
' At some point will also have to loop through active stories (EN. FN)
' Also `dictStyles` must be declared as global var.
  Set dictStyles = New genUtils.Dictionary
  Dim dictReturn As genUtils.Dictionary  ' the full dictionary object we'll return
  Dim dictInfo As genUtils.Dictionary   ' sub-sub dict for indiv. style info
  
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False
  dictReturn.Add "unstyledCount", 0    ' for now, just a count. can add more data later

  Dim lngParaCt As Long: lngParaCt = activeDoc.Paragraphs.Count
  Dim strStyle As String
  Dim strBodyStyle As String: strBodyStyle = "Text - Standard (tx)"
  Dim A As Long
  
' Loop through all paragraphs in document from END to START so we end up with
' FIRST page, and if we need to delete paras we don't mess up the count order
  For A = lngParaCt To 1 Step -1
  ' To break infinite loops.
  ' To do: increase? Add actual Err.Raise
    If A = 10000 Then
      Debug.Print "A = " & A
      Exit For
    End If
    
'    If A Mod 200 = 0 Then
'      Debug.Print "Paragraph " & A
'    End If

  ' Get style name
    strStyle = activeDoc.Paragraphs(A).Style
    
  ' If style name = Macmillan style...
    If Right(strStyle, 1) = ")" Then
    ' If style does not exist in dict yet...
      If Not dictStyles.Exists(strStyle) Then
      ' ...create sub-dictionary
        Set dictInfo = New genUtils.Dictionary
        dictInfo.Add "count", 0
        dictInfo.Add "startPara", 0
        dictStyles.Add strStyle, dictInfo
        Set dictInfo = Nothing
      End If
    ' Increase style count and update start paragraph index
    ' .Item() method overwrites value for that key
      dictStyles(strStyle).Item("count") = dictStyles(strStyle)("count") + 1
      dictStyles(strStyle).Item("startPara") = A
  ' Else (not Macmillan style)
    Else
    ' Increase unstyled count
      dictReturn.Item("unstyledCount") = dictReturn.Item("unstyledCount") + 1
      
    ' Change style, if requested
    ' To do: use logic to tag TX1, COTX1
    '        store style name externally
      If FixUnstyled = True Then
        activeDoc.Paragraphs(A).Style = strBodyStyle
      End If
    End If
  Next A
  
  ' What percentage are styled?
  Dim lngPercent As Single
  Dim blnPass As Boolean
  lngPercent = dictReturn("unstyledCount") / lngParaCt
  lngPercent = 1 - VBA.Round(lngPercent, 3)

' Threshold for "styled" is 50% of paragraphs have styles
  If lngPercent >= 0.5 Then
    blnPass = True
  Else
    blnPass = False
  End If
  
' update values in test dictionary
  dictReturn.Item("pass") = blnPass
  dictReturn.Item("uniqueStyles") = dictStyles.Count
  dictReturn.Item("percentStyled") = lngPercent
  
  Set StyleDictionary = dictReturn
  Exit Function

StyleDictionaryError:
  Err.Source = strReports & "StyleDictionary"
  If ErrorChecker(Err, strBodyStyle) = False Then
    Resume
  Else
    Call genUtils.Reports.ReportsTerminate
  End If
End Function


' ===== IsbnCheck =============================================================
' Call this to run ISBN checks.

Public Function IsbnCheck() As genUtils.Dictionary
  On Error GoTo IsbnCheckError
  Dim dictReturn As genUtils.Dictionary
  Set dictReturn = New Dictionary
  
  Dim strIsbnStyle As String
  strIsbnStyle = "span isbn (ISBN)"
  
' If no styled ISBN exists...
  Dim blnStyledIsbn As Boolean
  blnStyledIsbn = genUtils.GeneralHelpers.IsStyleInUse(strIsbnStyle)
  dictReturn.Add "styledIsbn", blnStyledIsbn
  If blnStyledIsbn = False Then
  
  ' Search for unstyled ISBN (returns tagged with bookmarks)
    Dim blnUnstyled As Boolean
    Dim arrISBN() As String
    arrISBN = IsbnSearch(ReturnArray:=True)
    If genUtils.IsArrayEmpty(arrISBN) = True Then
      blnUnstyled = False
    Else
      blnUnstyled = True
    End If
    dictReturn.Add "unstyledIsbn", blnUnstyled

  ' If no unstyled ISBNs, add from `book_info.json`, tag w/ bookmark
    If blnUnstyled = False Then
    ' If not found: Add Isbn
      Dim blnAddIsbn As Boolean
      blnAddIsbn = AddBookInfo(bk_ISBN)
      dictReturn.Add "isbnAdded", blnAddIsbn
    End If
    
  ' convert bookmarks to styles
    Dim bkName As Bookmark
    dictReturn.Add "taggedUnstyledIsbn", False
    For Each bkName In activeDoc.Bookmarks
      If Left(bkName.Name, 4) = "ISBN" Then
        bkName.Select
        Selection.Style = strIsbnStyle
        ' Report that we made a change
        dictReturn.Item("taggedUnstyled") = True
        bkName.Delete
      End If
    Next
  End If

' Cleanup what ISBN tag is covering
  Call genUtils.Reports.ISBNcleanup
  
' Read tagged isbns
  Dim isbnArray() As String   ' Even though they ARE numbers, keep as string
  isbnArray = genUtils.GeneralHelpers.GetText(strIsbnStyle, True)

' Add that this completed successfully?
  If genUtils.GeneralHelpers.IsArrayEmpty(isbnArray) = True Then
    dictReturn.Add "pass", True
    dictReturn.Add "list", isbnArray
  Else
    dictReturn.Add "pass", False
  End If
  
  Set IsbnCheck = dictReturn

  Exit Function
  
IsbnCheckError:
  Err.Source = strReports & "IsbnCheck"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.ReportsTerminate
  End If
End Function


' ===== IsbnSearch ============================================================
' Searches for unstyled ISBNs (13-digits with or without hyphens). If found,
' tags as bookmarks and returns array or string.

' PUBLIC because needs to be called independently from powershell if file name
' doesn't include ISBN. Optional FilePath is for passing doc path from PS.
' ReturnString is True by default also cuz powershell.

' Don't actually need to log anything with LogFile param, but powershell expects
' to pass that argument so we'll make it optional.

Public Function IsbnSearch(Optional FilePath As String, _
  Optional LogFile As String, _
  Optional ReturnString As Boolean = True) As Variant
  On Error GoTo IsbnSearchError
  Dim lngCounter As Long
  Dim strSearchPattern As String
  Dim ReturnArray() As String

  ' Make sure our document is open and active
  ' If not passing file path, will ref. activeDoc global var
  If FilePath <> vbNullString Then
    If genUtils.IsOpen(FilePath) = False Then
      Documents.Open (FilePath)
    End If
    Set activeDoc = Documents(FilePath)
  End If
  activeDoc.Range.Select
  ' ISBN rules:
  ' * First 3 digits: 978 or 979
  ' * 4th digit: 0 or 1 (for English language)
  ' * next section: publisher/imprint, 2 to 7 digits
  ' * next section: book, 1 to 6 digits (these two total 8 digits)
  ' * last section: single check digit
  ' sections may or may not be separated by a hyphen, but note that you can't use
  ' {0,1} to search for "zero or one" occurrences of something.
  ' the below is OK for now. Try more specific if needed later.
  strSearchPattern = "97[89][0-9\-]{10,14}"
  
  ' lngCounter both to prevent infinite loop and also for array index
  ' which we want to start at 0 because may pass back to powershell
  lngCounter = -1
  
  ' Start search at beginning of doc
  Selection.HomeKey Unit:=wdStory
  
  With Selection.Find
    .ClearFormatting
    .Text = strSearchPattern
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchWholeWord = False
    .MatchCase = True
    .MatchWildcards = True
    .MatchSoundsLike = False
  End With

  Do While Selection.Find.Execute = True And lngCounter < 100
    lngCounter = lngCounter + 1

    ' Delete if bookmark already exists
    If activeDoc.Bookmarks.Exists("ISBN" & lngCounter) = True Then
      activeDoc.Bookmarks.Item("ISBN" & lngCounter).Delete
    End If
    ' Add bookmark for later procedures to pick up
    activeDoc.Bookmarks.Add "ISBN" & lngCounter, Selection
    
    ' Also add to array to return to calling procedure
    ReDim Preserve ReturnArray(0 To lngCounter)
    ReturnArray(lngCounter) = Selection.Text
  Loop
  
  If ReturnString = False Then
    IsbnSearch = ReturnArray
  Else
    ' Default is comma-delimited string
    IsbnSearch = genUtils.Reduce(ReturnArray, ",")
  End If
  Exit Function

IsbnSearchError:
  Err.Source = strReports & "UnstyledIsbn"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.ReportsTerminate
  End If
End Function


' ===== AddBookInfo ===========================================================
' Add info from `book_info.json` to manuscript. Assume already know that it's
' not present. BookInfo is Private Enum.

Private Function AddBookInfo(InfoType As BookInfo) As Boolean
  On Error GoTo AddBookInfoError
  Dim strInfoKey As String
  Dim strInfoStyle As String
  Dim strInfoSection As String
  Dim strNewText As String
  
  ' Assign info key and ultimate paragraph style
  Select Case InfoType
    Case BookInfo.bk_Title
      strInfoKey = "title"
      strInfoStyle = "Titlepage Book Title (tit)"
    Case BookInfo.bk_Authors
      strInfoKey = "author"
      strInfoStyle = "Titlepage Author Name (au)"
    Case BookInfo.bk_ISBN
      strInfoKey = "isbn"
      strInfoStyle = "Copyright page single space (crtx)"
  End Select
  
' Get info string
  If dictBookInfo.Exists(strInfoKey) = True Then
    strNewText = dictBookInfo.Item(strInfoKey) & vbNewLine
  Else
    AddBookInfo = False
    Exit Function
  End If
  
' Find where this should go
  strInfoSection = Left(strInfoStyle, InStr(strInfoStyle, " ") - 1)
  
' Does section exist at all? Check in-use style dictionary for any style
  Dim key1 As Variant
  Dim lngCurrentStart As Long
  Dim lngStartPara As Long: lngStartPara = 1  ' Default if not found
  For Each key1 In dictStyles
    If InStr(key1, strInfoSection) > 0 Then
      lngCurrentStart = dictStyles(key1).Item("startPara")
      ' Should return LAST paragraph with that section's style.
      If lngCurrentStart > lngStartPara Then
        lngStartPara = lngCurrentStart
      End If
    End If
  Next key1
  
  ' Add text just before paragraph id'd above
  ' Once entered, new para takes index of lngStartPara.
  Dim rngNew As Range
  Set rngNew = activeDoc.Paragraphs(lngStartPara).Range
  rngNew.InsertBefore = strNewText
  rngNew.Style = strInfoStyle
  
  ' ISBN also needs character style
  If InfoType = bk_ISBN Then
    strInfoStyle = "span ISBN (isbn)"
    rngNew.Style = strInfoStyle
  End If
  
  Exit Function
AddBookInfoError:
  Err.Source = strReports & "AddBookInfo"
  If ErrorChecker(Err, strInfoStyle) = False Then
    Resume
  Else
    Call genUtils.ReportsTerminate
  End If
End Function


' ===== TitlepageCheck ========================================================
' Test that titlepage exists, Book Title exists, Author Name exists

Public Function TitlepageCheck() As genUtils.Dictionary
  On Error GoTo TitlepageCheckError
' set up return info
  Dim dictReturn As genUtils.Dictionary
  Set dictReturn = New genUtils.Dictionary
  With dictReturn
    .Add "pass", False
    .Add "bookTitleExists", False
    .Add "authorNameExists", False
  End With
  
' Some style names (eventually store externally?)
  Dim strTitle As String: strTitle = "Titlepage Book Title (tit)"
  Dim strAuthor As String: strAuthor = "Titlepage Author Name (au)"

  Dim blnTitle As Boolean
  Dim blnAuthor As Boolean

' Does Book Title exist?
  blnTitle = dictStyles.Exists(strBookTitle)
  dictReturn.Item("bookTitleExists") = blnTitle
  If blnTitle = False Then
    dictReturn.Item("bookTitleAdded") = AddBookInfo(bk_Title)
  End If

' Does Author Name exist?
  blnTitle = dictStyles.Exists(strAuthorName)
  dictReturn.Item("authorNameExists") = blnTitle
  If blnTitle = False Then
    dictReturn.Item("authorNameAdded") = AddBookInfo(bk_Authors)
  End If

' Did it all work?
  If genUtils.IsStyleInUse(strTitle) = True And _
    genUtils.IsStyleInUse(strAuthor) = True Then
    dictReturn.Item("pass") = True
  End If
  
  Set TitlepageCheck = dictReturn
  
  Exit Function
  
TitlepageCheckError:
  Err.Source strReports & "TitlepageCheck"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.Reports.ReportsTerminate
  End If
End Function


' ===== SectionCheck ==========================================================
' Tags book sections, adds to dictionary as ranges. Also fixes breaks.

Public Function SectionCheck() As genUtils.Dictionary
  On Error GoTo SectionCheckError
  ' Create array of paragrpah indices of page breaks
  
  ' Create array of section ranges based on those page breaks
  Exit Function
SectionCheckError:
  Err.Source = strReports & "SectionCheck"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.Reports.ReportsTerminate
  End If
End Function

' ===== StyleCleanup ==========================================================
' Tweaking some styles that will cause problems. Can probably cut once we
' update style name list.

Private Function StyleCleanup() As genUtils.Dictionary
  On Error GoTo StyleCleanupError
  Dim dictReturn As genUtils.Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False
' Change "FM Epigraph" to just "Epigraph" so we can determine section
' Hard-code styles cuz we only need this until we change style names
  Dim strFmEpis(1 To 3) As String
  Dim X As Long
  strFmEpis(1) = "FM Epigraph - non-verse (fmepi)"
  strFmEpis(2) = "FM Epigraph - verse (fmepiv)"
  strFmEpis(3) = "FM Epigraph Source (fmeps)"

' Loop through to see if these are even in use. If yes, replace.
  Dim strNewEpi As String
  For X = LBound(strFmEpis) To UBound(strFmEpis)
    If genUtils.GeneralHelpers.IsStyleInUse(strFmEpis(X)) = True Then
      dictReturn.Add "convertFmEpi", True
    ' Convert to correct style name (vbTextCompare = case insensitive)
      strNewEpi = VBA.LTrim(VBA.Replace(strFmEpis(X), "FM", "", _
        Compare:=vbTextCompare))
    ' Find/Replace all instances. Will error if new style not present, but
    ' ErrorChecker will add if we pass the new style name to it.
      With activeDoc.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Format = True
        .Style = strFmEpis(X)
        .Replacement.Style = strNewEpi
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
      End With
    End If
  Next X

' Remove any section break characters. Can't assume they'll be in their own
' paragraphs, so remove solo paragraphs in next step.
  With activeDoc.Range.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = "^b"
    .Replacement.Text = ""
    .Format = True
    .Style = strFmEpis(X)
    .Replacement.Style = strNewEpi
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .Execute Replace:=wdReplaceAll
  
    If .Found = True Then
      dictReturn.Add "deleteSectionBrk", True
    Else
      dictReturn.Add "deleteSectionBrk", False
    End If
  End With
  

' Convert any Section Break styles to Page Break (because they may have been
' added after a Page Break and we don't want to confuse PageBreak fixes later
' Use variable for new style name in case it's not present
  strNewEpi = "Page Break (pb)"
  With activeDoc.Range.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = ""
    .Replacement.Text = ""
    .Format = True
    .Style = "Section Break (sbr)"
    .Replacement.Style = strNewEpi
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .Execute Replace:=wdReplaceAll
  
    If .Found = True Then
      dictReturn.Add "deleteSectionStyle", True
    Else
      dictReturn.Add "deleteSectionStyle", False
    End If
  End With

' Remove any Half Title paras. (If want to keep in future, create a separate
' function to search for all half titles, add headings/breaks. Note that any
' extra page breaks will get cleaned up in `PageBreakCleanup` function.
  strNewEpi = "Halftitle Book Title (htit)"
  With activeDoc.Range.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = "*"
    .Replacement.Text = ""
    .Format = True
    .Style = strNewEpi
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
  
    If .Found = True Then
      dictReturn.Add "deleteHalfTitle", True
    Else
      dictReturn.Add "deleteHalfTitle", False
    End If
  End With

  Call genUtils.zz_clearFind
  
  dictReturn.Item("pass") = True
  Set StyleCleanup = dictReturn
  Exit Function

StyleCleanupError:
  Err.Source = strReports & "StyleCleanup"
  If ErrorChecker(Err, strNewEpi) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function

' ===== IsHeading =============================================================
' Is this paragraph style a Macmillan heading style? Eventually store style
' names externally.

Private Function IsHeading(StyleName As String) As Boolean
  On Error GoTo IsHeadingError
' Hard code for now. `dictHeadings` is global scope so only have to create once
  If dictHeadings Is Nothing Then
    Set dictHeadings = New Dictionary
    With dictHeadings
      .Add "Front Sales Quote Head (fsqh)"
      .Add "FM Head (fmh)"
      .Add "FM Title (fmt)"
      .Add "Fm Head ALT (afmh)"
      .Add "Chap Title (ct)"
      .Add "Chap Number (cn)"
      .Add "Chap Title Nonprinting (ctnp)"
      .Add "Part Title (pt)"
      .Add "Part Number (pn)"
      .Add "BM Head (bmh)"
      .Add "BM Title (bmt)"
      .Add "BM Head ALT (abmh)"
      .Add "Appendix Head (aph)"
      .Add "About Author Text Head (atah)"
      .Add "Series Page Heading (sh)"
      .Add "Ad Card Main Head (acmh)"
      .Add "Recipe Head (rh)"
      .Add "Sub-Recipe Head (srh)"
      .Add "Recipe Var Head (rvh)"
      .Add "Poem Subtitle (vst)"
      .Add "Poem Title (vt)"
    End With
  End If
  
' So just see if our style is one of these styles
  IsHeading = dictHeadings.Exists(StyleName)
  Exit Function
IsHeadingError:
  Err.Source = strReports & "IsHeading"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Cell genUtils.ReportsTerminate
  End If
End Function


' ===== SectionName ===========================================================
' Determines section name from style name. Reads from external JSON file. When
' we get around to changing style names so the section is always the first word
' we can make this much simpler.

Private Function SectionName(StyleName As String) As String
  On Error GoTo SectionNameError
' Create dictionary if it hasn't been yet
  If dictSections Is Nothing Then
  ' Check for `sections.json` file, read into global dictionary
    Dim strSections As String
    strSections = ThisDocument.Path & Application.PathSeparator & "sections.json"
    If genUtils.IsItThere(strSections) = True Then
      Set dictSections = genUtils.ClassHelpers.ReadJson(strSections)
    Else
      Err.Raise MacError.err_FileNotThere
    End If
  End If

' If first word is a key, section is value
  Dim strFirst As String
  strFirst = Left(StyleName, InStr(StyleName, " ") - 1)
  If dictSections.Exists(strFirst) = True Then
    SectionName = dictSections.Item(strFirst)
' Else it's just a generic chapter section
  Else
    SectionName = "Chapter"
  End If
  Exit Function

SectionNameError:
  Err.Source = strReports & "SectionName"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    genUtils.ReportsTerminate
  End If
End Function


' ===== AddHeading ============================================================
' Adds CTNP heading ABOVE the paragraph passed as arg, with section name text.
' Though each chapter is just "Chapter" -- will have to add numbers later.

Private Function AddHeading(ParaInd As Long) As Boolean
  On Error GoTo AddHeadingError
' Get current para count to test at end that we added a new one
  Dim lngParas As Long
  lngParas = activeDoc.Paragraphs.Count

  ' generic head style (eventually store externally)
  Dim strHeadStyle As String
  strHeadStyle = "Chap Title Nonprinting (ctnp)"
  
' Set range for the para in question
  Dim rngPara As Range
  Set rngPara = activeDoc.Paragraphs(ParaInd).Range
  
' Get style name of that paragraph
  Dim strParaStyle As String
  strParaStyle = rngPara.Style

' Look up section name of that style
  Dim strSectionName As String
  strSectionName = Reports.SectionName(strParaStyle)
  ' add line ending ('cuz new paragraph)
  strSectionName = strSectionName & vbNewLine
  
' Insert new paragraph
  rngPara.InsertBefore (strSectionName)
  
' Add correct style (inserted paragraph now part of `rngPara` object)
  rngPara.Paragraphs(1) = strHeadStyle

' Verify we added a paragraph
  Dim lngNewParas As Long
  lngNewParas = activeDoc.Paragraphs.Count
  If lngNewParas = lngParas + 1 Then
    AddHeading = True
  Else
    AddHeading = False
  End If
  Exit Function
  
AddHeadingError:
  Err.Source strReports & "AddHeading"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.ReportsTerminate
  End If
End Function

' ===== PageBreakCleanup ======================================================
' Clean up page break characters/styles, so just single paragraph break chars
' styles as "Page Break" remain.

Private Function PageBreakCleanup() As Boolean
  On Error GoTo PageBreakCleanupError

' Add paragraph breaks around every page break character (so we know for sure
' paragraph style of break won't apply to any body text). Will add extra blank
' paragraphs that we can clean up later.


' Apply "Page Break (pb)" style to every PB char (to catch any unstyled PB)


' Now that we are sure every PB char has PB style, remove all PB char


' Remove multiple PB-styled paragraphs in a row


  Exit Function

PageBreakCleanupError:
  Err.Source = strReports & "PageBreakCleanup"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function

' ===== PageBreakCheck ========================================================
' Check that every page break is followed by a heading. If not, add one.

Private Function PageBreakCheck() As Boolean
  On Error GoTo PageBreakCheckError

' Loop through search of all "Page Break"-styled paragraphs


' If the NEXT paragraph is NOT an approved heading style...


' ... determine the section and add a CTNP heading


  Exit Function
  
PageBreakCheckError:
  Err.Source = strReports & "PageBreakCheck"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.ReportsTerminate
  End If
End Function
' #############################################################################
' =============================================================================
'
'       OLD REPORTS CODE BELOW
'
' =============================================================================
' #############################################################################

'Private Function GoodBadStyles(Tor As Boolean, ProgressBar As ProgressBar, _
'Status As String, ProgTitle As String, Stories() As Variant) As Variant
'    'Creates a list of Macmillan styles in use
'    'And a separate list of non-Macmillan styles in use
'
'    Dim TheOS As String
'    TheOS = System.OperatingSystem
'    Dim sglPercentComplete As Single
'    Dim strStatus As String
'
'    Dim activeDoc As Document
'    Set activeDoc = ActiveDocument
'    Dim stylesGood() As String
'    Dim stylesGoodLong As Long
'    stylesGoodLong = 400                                    'could maybe reduce this number
'    ReDim stylesGood(stylesGoodLong)
'    Dim stylesBad() As String
'    ReDim stylesBad(1 To 100) 'could maybe reduce this number too
'    Dim styleGoodCount As Integer
'    Dim styleBadCount As Integer
'    Dim styleBadOverflow As Boolean
'    Dim activeParaCount As Integer
'    Dim J As Integer, K As Integer, L As Integer
'    Dim paraStyle As String
'    '''''''''''''''''''''
'    Dim activeParaRange As Range
'    Dim pageNumber As Integer
'    Dim A As Long
'
'
'
'
'    '----------Collect all styles being used-------------------------------
'    styleGoodCount = 0
'    styleBadCount = 0
'    styleBadOverflow = False
'    activeParaCount = activeDoc.Paragraphs.Count
'    For J = 1 To activeParaCount
'
'        'All Progress Bar statements for PC only because won't run modeless on Mac
'        If J Mod 100 = 0 Then
'
'            'Percent complete and status for progress bar (PC) and status bar (Mac)
'            sglPercentComplete = (((J / activeParaCount) * 0.45) + 0.18)
'            strStatus = "* Checking paragraph " & J & " of " & activeParaCount & " for Macmillan styles..." & _
'                        vbCr & Status
'
'            'Debug.Print sglPercentComplete
'            Call UpdateBarAndWait(Bar:=ProgressBar, Status:=strStatus, Percent:=sglPercentComplete)
'        End If
'
'        For A = LBound(Stories()) To UBound(Stories())
'            If J <= ActiveDocument.StoryRanges(Stories(A)).Paragraphs.Count Then
'                paraStyle = activeDoc.StoryRanges(Stories(A)).Paragraphs(J).Style
'                Set activeParaRange = activeDoc.StoryRanges(Stories(A)).Paragraphs(J).Range
'                pageNumber = activeParaRange.Information(wdActiveEndPageNumber)                 'alt: (wdActiveEndAdjustedPageNumber)
'
'                'If InStrRev(paraStyle, ")", -1, vbTextCompare) Then        'ALT calculation to "Right", can speed test
'                If Right(paraStyle, 1) = ")" Then
'CheckGoodStyles:
'                    For K = 1 To styleGoodCount
'                        'Debug.Print Left(stylesGood(K), InStrRev(stylesGood(K), " --") - 1)
'                        ' "Left" function because now stylesGood includes page number, so won't match paraStyle
'                        If paraStyle = Left(stylesGood(K), InStrRev(stylesGood(K), " --") - 1) Then
'                        K = styleGoodCount                              'stylereport bug fix #1    v. 3.1
'                            Exit For                                        'stylereport bug fix #1   v. 3.1
'                        End If                                              'stylereport bug fix #1   v. 3.1
'                    Next K
'
'                    If K = styleGoodCount + 1 Then
'                        styleGoodCount = K
'                        ReDim Preserve stylesGood(1 To styleGoodCount)
'                        stylesGood(styleGoodCount) = paraStyle & " -- p. " & pageNumber
'                    End If
'
'                Else
'
'                    If paraStyle = "Endnote Text" Or paraStyle = "Footnote Text" Then
'                        GoTo CheckGoodStyles
'                    Else
'                        For L = 1 To styleBadCount
'                            'If paraStyle = stylesBad(L) Then Exit For                  'Not needed, since we want EVERY instance of bad style
'                        Next L
'                        If L > 100 Then                                                 ' Exits if more than 100 bad paragraphs
'                                styleBadOverflow = True
'                                stylesBad(100) = "** WARNING: More than 100 paragraphs with bad styles found." & vbNewLine & vbNewLine
'                            Exit For
'                        End If
'                        If L = styleBadCount + 1 Then
'                            styleBadCount = L
'
'                            stylesBad(styleBadCount) = "** ERROR: Non-Macmillan style on page " & pageNumber & _
'                                " (Paragraph " & J & "):  " & paraStyle & vbNewLine & vbNewLine
'                        End If
'                     End If
'                End If
'            End If
'        Next A
'    Next J
'
'    Status = "* Checking paragraphs for Macmillan styles..." & vbCr & Status
'
'    'Change Normal (Web) back (if you want to)
'    ActiveDocument.Styles("Normal (Web),_").NameLocal = "Normal (Web)"
'
'    ' DON'T sort styles alphabetically, per request from PE
''    'Sort good styles
''    If K <> 0 Then
''    ReDim Preserve stylesGood(1 To styleGoodCount)
''    WordBasic.SortArray stylesGood()
''    End If
'
'    'Create single string for good styles
'    Dim strGoodStyles As String
'
'    If styleGoodCount = 0 Then
'        strGoodStyles = ""
'    Else
'        For K = LBound(stylesGood()) To UBound(stylesGood())
'            strGoodStyles = strGoodStyles & stylesGood(K) & vbCrLf
'        Next K
'    End If
'
'    'Debug.Print strGoodStyles
'
'    If styleBadCount > 0 Then
'        'Create single string for bad styles
'        Dim strBadStyles As String
'        ReDim Preserve stylesBad(1 To styleBadCount)
'        For L = LBound(stylesBad()) To UBound(stylesBad())
'            strBadStyles = strBadStyles & stylesBad(L)
'        Next L
'    Else
'        strBadStyles = ""
'    End If
'
'    'Debug.Print strBadStyles
'
'    '-------------------get list of good character styles--------------
'
'    Dim charStyles As String
'    Dim styleNameM(1 To 21) As String        'declare number in array
'    Dim M As Integer
'
'    styleNameM(1) = "span italic characters (ital)"
'    styleNameM(2) = "span boldface characters (bf)"
'    styleNameM(3) = "span small caps characters (sc)"
'    styleNameM(4) = "span underscore characters (us)"
'    styleNameM(5) = "span superscript characters (sup)"
'    styleNameM(6) = "span subscript characters (sub)"
'    styleNameM(7) = "span bold ital (bem)"
'    styleNameM(8) = "span smcap ital (scital)"
'    styleNameM(9) = "span smcap bold (scbold)"
'    styleNameM(10) = "span symbols (sym)"
'    styleNameM(11) = "span accent characters (acc)"
'    styleNameM(12) = "span cross-reference (xref)"
'    styleNameM(13) = "span hyperlink (url)"
'    styleNameM(14) = "span material to come (tk)"
'    styleNameM(15) = "span carry query (cq)"
'    styleNameM(16) = "span preserve characters (pre)"
'    styleNameM(17) = "span strikethrough characters (str)"
'    styleNameM(18) = "bookmaker keep together (kt)"
'    styleNameM(19) = "span ISBN (isbn)"
'    styleNameM(20) = "span symbols ital (symi)"
'    styleNameM(21) = "span symbols bold (symb)"
'
'
'
'    For M = 1 To UBound(styleNameM())
'
'        'Percent complete and status for progress bar (PC) and status bar (Mac)
'        sglPercentComplete = (((M / UBound(styleNameM())) * 0.13) + 0.63)
'        strStatus = "* Checking for " & styleNameM(M) & " styles..." & vbCr & Status
'
'        Call UpdateBarAndWait(Bar:=ProgressBar, Status:=strStatus, Percent:=sglPercentComplete)
'
'        On Error GoTo ErrHandler
'
'        'Move selection back to start of document
'        Selection.HomeKey Unit:=wdStory
'
'        'Need to do Selection.Find for char styles. Range.Find won't work.
'        'We only need to find a style once to add it to the list
'        'Search through the main text story here
'        With Selection.Find
'            .Style = ActiveDocument.Styles(styleNameM(M))
'            .Wrap = wdFindContinue
'            .Format = True
'            .Execute
'        End With
'
'        If Selection.Find.Found = True Then
'            charStyles = charStyles & styleNameM(M) & vbNewLine
'        'Else not present in main text story
'        Else
'            ' So check if there are footnotes
'            If ActiveDocument.Footnotes.Count > 0 Then
'                'If there are footnotes, select the footnote text
'                ActiveDocument.StoryRanges(wdFootnotesStory).Select
'                'Search the new selection for the style
'                With Selection.Find
'                    .Style = ActiveDocument.Styles(styleNameM(M))
'                    .Wrap = wdFindContinue
'                    .Format = True
'                    .Execute
'                End With
'
'                If Selection.Find.Found = True Then
'                    charStyles = charStyles & styleNameM(M) & vbNewLine
'                ' Else didn't find style in footnotes, check endnotes
'                Else
'                    GoTo CheckEndnotes
'                End If
'            Else
'CheckEndnotes:
'                ' Check if there are endnotes in the document
'                If ActiveDocument.Endnotes.Count > 0 Then
'                    ' If there are endnotes, select them
'                    ActiveDocument.StoryRanges(wdEndnotesStory).Select
'                    'Search the new selection for the style
'                    With Selection.Find
'                         .Style = ActiveDocument.Styles(styleNameM(M))
'                         .Wrap = wdFindContinue
'                         .Format = True
'                         .Execute
'                     End With
'
'                    If Selection.Find.Found = True Then
'                        charStyles = charStyles & styleNameM(M) & vbNewLine
'                    End If
'                End If
'            End If
'        End If
'NextLoop:
'    Next M
'
'    'Debug.Print charStyles
'
'    Status = "* Checking character styles..." & vbCr & Status
'
'    'Add character styles to Good styles list
'    strGoodStyles = strGoodStyles & charStyles
'
'    'If this is for the Tor.com Bookmaker toolchain, test if only those styles used
'    Dim strTorBadStyles As String
'    If Tor = True Then
'        strTorBadStyles = BadTorStyles(ProgressBar2:=ProgressBar, StatusBar:=Status, ProgressTitle:=ProgTitle, Stories:=Stories)
'        strBadStyles = strBadStyles & strTorBadStyles
'    End If
'
'    'Debug.Print strGoodStyles
'    'Debug.Print strBadStyles
'
'    'If only good styles are Endnote Text and Footnote text, then the template is not being used
'
'
'    'Add both good and bad styles lists to an array to pass back to original sub
'    Dim arrFinalLists() As Variant
'    ReDim arrFinalLists(1 To 2)
'
'    arrFinalLists(1) = strGoodStyles
'    arrFinalLists(2) = strBadStyles
'
'    GoodBadStyles = arrFinalLists
'
'    Exit Function
'
'ErrHandler:
'    'Debug.Print Err.Number & " : " & Err.Description
'    If Err.Number = 5834 Or Err.Number = 5941 Then
'        Resume NextLoop
'    End If
'
'End Function
'
'
'Private Function CreateErrorList(badStyles As String, arrStyleCount() As Variant, blnTor As Boolean) As String
'    Dim errorList As String
'
'    errorList = ""
'
'    '--------------For reference----------------------
'    'arrStyleCount(1) = "Titlepage Book Title (tit)"
'    'arrStyleCount(2) = "Titlepage Author Name (au)"
'    'arrStyleCount(3) = "span ISBN (isbn)"
'    'arrStyleCount(4) = "Chap Number (cn)"
'    'arrStyleCount(5) = "Chap Title (ct)"
'    'arrStyleCount(6) = "Chap Title Nonprinting (ctnp)"
'    'arrStyleCount(7) = "Titlepage Logo (logo)"
'    'arrStyleCount(8) = "Part Title (pt)"
'    'arrStyleCount(9) = "Part Number (pn)"
'    'arrStyleCount(10) = "FM Head (fmh)"
'    'arrStyleCount(11) = "FM Title (fmt)"
'    'arrStyleCount(12) = "BM Head (bmh)"
'    'arrStyleCount(13) = "BM Title (bmt)"
'    'arrStyleCount(14) = "Illustration holder (ill)"
'    'arrStyleCount(15) = "Illustration source (is)"
'    '------------------------------------------------
'
'    '=====================Generate errors based on number of required elements found==================
'
'    'If Book Title = 0
'    If arrStyleCount(1) = 0 Then errorList = errorList & "** ERROR: No styled title detected." & _
'        vbNewLine & vbNewLine
'
'    'If Book Title > 1
'    If arrStyleCount(1) > 1 Then errorList = errorList & "** ERROR: Too many title paragraphs detected." _
'        & " Only 1 allowed." & vbNewLine & vbNewLine
'
'    'Check if page break before Book Title
'    If arrStyleCount(1) > 0 Then errorList = errorList & CheckPrevStyle(findStyle:="Titlepage Book Title (tit)", _
'        prevStyle:="Page Break (pb)")
'
'
'    'If Author Name = 0
'    If arrStyleCount(2) = 0 Then errorList = errorList & "** ERROR: No styled author name detected." _
'        & vbNewLine & vbNewLine
'
'    'If ISBN = 0
'    If arrStyleCount(3) = 0 Then
'        errorList = errorList & "** ERROR: No styled ISBN detected." _
'        & vbNewLine & vbNewLine
'    Else
'        If blnTor = True Then
'            'check for correct book type following ISBN, in parens.
'            errorList = errorList & BookTypeCheck
'        End If
'    End If
'
'    'If CN > 0 and CT = 0 (already fixed in FixSectionHeadings sub)
'    If arrStyleCount(4) > 0 And arrStyleCount(5) = 0 Then errorList = errorList & _
'        "** WARNING: Chap Number (cn) cannot be the main heading for" & vbNewLine _
'        & vbTab & "a chapter. Every chapter must include Chapter Title (ct)" & vbNewLine _
'        & vbTab & "style. Chap Number (cn) paragraphs have been converted to the" & vbNewLine _
'        & vbTab & "Chap Title (ct) style." & vbNewLine & vbNewLine
'
'    'If PN > 0 and PT = 0 (already fixed in FixSectionHeadings sub)
'    If arrStyleCount(9) > 0 And arrStyleCount(8) = 0 Then errorList = errorList & _
'        "** WARNING: Part Number (pn) cannot be the main heading for" & vbNewLine _
'        & vbTab & "a section. Every part must include Part Title (pt)" & vbNewLine _
'        & vbTab & "style. Part Number (pn) paragraphs have been converted" & vbNewLine _
'        & vbTab & "to the Part Title (pt) style." & vbNewLine & vbNewLine
'
'    'If FMT > 0 and FMH = 0 (already fixed in FixSectionHeadings sub)
'    If arrStyleCount(11) > 0 And arrStyleCount(10) = 0 Then errorList = errorList & _
'        "** WARNING: FM Title (fmt) cannot be the main heading for" & vbNewLine _
'        & vbTab & "a section. Every front matter section must include" & vbNewLine _
'        & vbTab & "the FM Head (fmh) style. FM Title (fmt) paragraphs" & vbNewLine _
'        & vbTab & "have been converted to the FM Head (fmh) style." & vbNewLine & vbNewLine
'
'    'If BMT > 0 and BMH = 0 (already fixed in FixSectionHeadings sub)
'    If arrStyleCount(13) > 0 And arrStyleCount(12) = 0 Then errorList = errorList & _
'        "** WARNING: BM Title (bmt) cannot be the main heading for" & vbNewLine _
'        & vbTab & "a section. Every back matter section must incldue" & vbNewLine _
'        & vbTab & "the BM Head (bmh) style. BM Title (bmt) paragraphs" & vbNewLine _
'        & vbTab & "have been converted to the BM Head (bmh) style." & vbNewLine & vbNewLine
'
'    'If no chapter opening paragraphs (CN, CT, or CTNP)
'    If arrStyleCount(4) = 0 And arrStyleCount(5) = 0 And arrStyleCount(6) = 0 Then errorList = errorList _
'        & "** ERROR: No tagged chapter openers detected. If your book does" & vbNewLine _
'        & vbTab & "not have chapter openers, use the Chap Title Nonprinting" & vbNewLine _
'        & vbTab & "(ctnp) style at the start of each section." & vbNewLine & vbNewLine
'
'    'If CN > CT and CT > 0 (i.e., Not a CT for every CN)
'    If arrStyleCount(4) > arrStyleCount(5) And arrStyleCount(5) > 0 Then errorList = errorList & _
'        "** ERROR: More Chap Number (cn) paragraphs than Chap Title (ct)" & vbNewLine _
'        & vbTab & "paragraphs found. Each Chap Number (cn) paragraph MUST be" & vbNewLine _
'        & vbTab & "followed by a Chap Title (ct) paragraph." & vbNewLine & vbNewLine
'
'    'If Imprint line = 0
'    If arrStyleCount(7) = 0 Then errorList = errorList & "** WARNING: No styled Titlepage Logo (logo) line detected. " _
'        & "If you would like a logo included on your titlepage, please add this style." & vbNewLine & vbNewLine
'
'    'If Imprint Lline > 1
'    If arrStyleCount(7) > 1 Then errorList = errorList & "** ERROR: Too many Imprint Line paragraphs" _
'        & " detected. Only 1 allowed." & vbNewLine & vbNewLine
'
'    'If only CTs because converted by macro check for a page break before
'    If (arrStyleCount(4) > 0 And arrStyleCount(5) = 0) Then errorList = errorList & _
'        CheckPrevStyle(findStyle:="Chap Title (ct)", prevStyle:="Page Break (pb)")
'
'    'If only PTs (either originally or converted by macro) check for a page break before
'    If (arrStyleCount(9) > 0 And arrStyleCount(8) = 0) Or (arrStyleCount(9) = 0 And arrStyleCount(8) > 0) _
'        Then errorList = errorList & CheckPrevStyle(findStyle:="Part Title (pt)", prevStyle:="Page Break (pb)")
'
'    'If only FMHs (either originally or converted by macro) check for a page break before
'    If (arrStyleCount(11) > 0 And arrStyleCount(10) = 0) Or (arrStyleCount(11) = 0 And arrStyleCount(10) > 0) _
'        Then errorList = errorList & CheckPrevStyle(findStyle:="FM Head (fmh)", prevStyle:="Page Break (pb)")
'
'    'If only BMHs (either originally or converted by macro) check for a page break before
'    If (arrStyleCount(13) > 0 And arrStyleCount(12) = 0) Or (arrStyleCount(13) = 0 And arrStyleCount(12) > 0) _
'        Then errorList = errorList & CheckPrevStyle(findStyle:="BM Head (bmh)", prevStyle:="Page Break (pb)")
'
'    'If only CTP, check for a page break before
'    If arrStyleCount(4) = 0 And arrStyleCount(5) = 0 And arrStyleCount(6) > 0 Then errorList = errorList _
'        & CheckPrevStyle(findStyle:="Chap Title Nonprinting (ctnp)", prevStyle:="Page Break (pb)")
'
'    'If CNs <= CTs, then check that those 3 styles are in order
'    If arrStyleCount(4) <= arrStyleCount(5) And arrStyleCount(4) > 0 Then errorList = errorList & CheckPrev2Paras("Page Break (pb)", _
'        "Chap Number (cn)", "Chap Title (ct)")
'
'    'If Illustrations and sources exist, check that source comes after Ill and Cap
'    If blnTor = True Then
'        If arrStyleCount(14) > 0 And arrStyleCount(15) > 0 Then errorList = errorList & _
'            CheckPrev2Paras("Illustration holder (ill)", "Caption (cap)", "Illustration Source (is)")
'        If CheckFileName = True Then errorList = errorList & _
'            "**ERROR: Bookmaker can only accept file names that use" & vbNewLine & _
'            "letters, numbers, hyphens, or underscores. Punctuation," & vbNewLine & _
'            "spaces, and other special characters are not allowed." & vbNewLine & vbNewLine
'    End If
'
'    'Check that only heading styles follow page breaks
'    errorList = errorList & CheckAfterPB
'
'    ' Check that all CTNP have some text
'    If arrStyleCount(6) > 0 Then errorList = errorList & CheckNonprintingText
'
'    'Add bad styles to error message
'    errorList = errorList & badStyles
'
'    If errorList <> "" Then
'        errorList = errorList & vbNewLine & "If you have any questions about how to handle these errors, " & vbNewLine & _
'            "please contact workflows@macmillan.com." & vbNewLine
'    End If
'
'    'Debug.Print errorList
'
'    CreateErrorList = errorList
'
'End Function
'
'
'Function CheckPrevStyle(findStyle As String, prevStyle As String) As String
'    Dim jString As String
'    Dim jCount As Integer
'    Dim pageNum As Integer
'    Dim intCurrentPara As Integer
'
'    Application.ScreenUpdating = False
'
'        'check if styles exist, else exit sub
'        On Error GoTo ErrHandler:
'        Dim keyStyle As Word.Style
'
'        Set keyStyle = ActiveDocument.Styles(findStyle)
'        Set keyStyle = ActiveDocument.Styles(prevStyle)
'
'    jCount = 0
'    jString = ""
'
'    'Move selection to start of document
'    Selection.HomeKey Unit:=wdStory
'
'    'select paragraph with that style
'        Selection.Find.ClearFormatting
'        With Selection.Find
'            .Text = ""
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindStop
'            .Format = True
'            .Style = ActiveDocument.Styles(findStyle)
'            .MatchCase = False
'            .MatchWholeWord = False
'            .MatchWildcards = False
'            .MatchSoundsLike = False
'            .MatchAllWordForms = False
'        End With
'
'    Do While Selection.Find.Execute = True And jCount < 200            'jCount so we don't get an infinite loop
'        jCount = jCount + 1
'
'        'Get number of current pagaraph, because we get an error if try to select before 1st para
'        Dim rParagraphs As Range
'        Dim CurPos As Long
'
'        Selection.Range.Select  'select current ran
'        CurPos = ActiveDocument.Bookmarks("\startOfSel").Start
'        Set rParagraphs = ActiveDocument.Range(Start:=0, End:=CurPos)
'        intCurrentPara = rParagraphs.Paragraphs.Count
'
'        'Debug.Print intCurrentPara
'
'        If intCurrentPara > 1 Then
'            'select preceding paragraph
'            Selection.Previous(Unit:=wdParagraph, Count:=1).Select
'            pageNum = Selection.Information(wdActiveEndPageNumber)
'
'                'Check if preceding paragraph style is correct
'                If Selection.Style <> prevStyle Then
'                    jString = jString & "** ERROR: Missing or incorrect " & prevStyle & " style before " _
'                        & findStyle & " style on page " & pageNum & "." & vbNewLine & vbNewLine
'                End If
'
'                'If you're searching for a page break before, also check if manual page break is in paragraph
'                If prevStyle = "Page Break (pb)" Then
'                    If InStr(Selection.Text, Chr(12)) = 0 Then
'                        jString = jString & "** ERROR: Missing manual page break on page " & pageNum & "." _
'                            & vbNewLine & vbNewLine
'                    End If
'                End If
'
'                'Debug.Print jString
'
'            'move the selection back to original paragraph, so it won't be
'            'selected again on next search
'            Selection.Next(Unit:=wdParagraph, Count:=1).Select
'        End If
'
'    Loop
'
'    'Debug.Print jString
'
'    CheckPrevStyle = jString
'
'    Exit Function
'
'ErrHandler:
'    If Err.Number = 5941 Or Err.Number = 5834 Then       'style doesn't exist in document
'        Exit Function
'    End If
'End Function
'
'Function CheckAfterPB() As String
'    Dim arrSecStartStyles() As String
'    ReDim arrSecStartStyles(1 To 44)
'    Dim kString As String
'    Dim kCount As Integer
'    Dim pageNumK As Integer
'    Dim nextStyle As String
'    Dim N As Integer
'    Dim nCount As Integer
'
'    Application.ScreenUpdating = False
'
'    ' These are all styles allowed to follow a page break
'    arrSecStartStyles(1) = "Chap Title (ct)"
'    arrSecStartStyles(2) = "Chap Number (cn)"
'    arrSecStartStyles(3) = "Chap Title Nonprinting (ctnp)"
'    arrSecStartStyles(4) = "Halftitle Book Title (htit)"
'    arrSecStartStyles(5) = "Titlepage Book Title (tit)"
'    arrSecStartStyles(6) = "Copyright Text single space (crtx)"
'    arrSecStartStyles(7) = "Copyright Text double space (crtxd)"
'    arrSecStartStyles(8) = "Dedication (ded)"
'    arrSecStartStyles(9) = "Ad Card Main Head (acmh)"
'    arrSecStartStyles(10) = "Ad Card List of Titles (acl)"
'    arrSecStartStyles(11) = "Part Title (pt)"
'    arrSecStartStyles(12) = "Part Number (pn)"
'    arrSecStartStyles(13) = "Front Sales Title (fst)"
'    arrSecStartStyles(14) = "Front Sales Quote (fsq)"
'    arrSecStartStyles(15) = "Front Sales Quote NoIndent (fsq1)"
'    arrSecStartStyles(16) = "Epigraph - non-verse (epi)"
'    arrSecStartStyles(17) = "Epigraph - verse (epiv)"
'    arrSecStartStyles(18) = "FM Head (fmh)"
'    arrSecStartStyles(19) = "Illustration holder (ill)"
'    arrSecStartStyles(20) = "Page Break (pb)"
'    arrSecStartStyles(21) = "FM Epigraph - non-verse (fmepi)"
'    arrSecStartStyles(22) = "FM Epigraph - verse (fmepiv)"
'    arrSecStartStyles(23) = "FM Head ALT (afmh)"
'    arrSecStartStyles(24) = "Part Epigraph - non-verse (pepi)"
'    arrSecStartStyles(25) = "Part Epigraph - verse (pepiv)"
'    arrSecStartStyles(26) = "Part Contents Main Head (pcmh)"
'    arrSecStartStyles(27) = "Poem Title (vt)"
'    arrSecStartStyles(28) = "Recipe Head (rh)"
'    arrSecStartStyles(29) = "Sub-Recipe Head (srh)"
'    arrSecStartStyles(30) = "BM Head (bmh)"
'    arrSecStartStyles(31) = "BM Head ALT (abmh)"
'    arrSecStartStyles(32) = "Appendix Head (aph)"
'    arrSecStartStyles(33) = "About Author Text (atatx)"
'    arrSecStartStyles(34) = "About Author Text No-Indent (atatx1)"
'    arrSecStartStyles(35) = "About Author Text Head (atah)"
'    arrSecStartStyles(36) = "Colophon Text (coltx)"
'    arrSecStartStyles(37) = "Colophon Text No-Indent (coltx1)"
'    arrSecStartStyles(38) = "BOB Ad Title (bobt)"
'    arrSecStartStyles(39) = "Series Page Heading (sh)"
'    arrSecStartStyles(40) = "span small caps characters (sc)"
'    arrSecStartStyles(41) = "span italic characters (ital)"
'    arrSecStartStyles(42) = "Design Note (dn)"
'    arrSecStartStyles(43) = "Front Sales Quote Head (fsqh)"
'    arrSecStartStyles(44) = "Section Break (sbr)"
'
'    kCount = 0
'    kString = ""
'
'    'Move selection to start of document
'    Selection.HomeKey Unit:=wdStory
'
'    On Error GoTo ErrHandler1
'
'    'select paragraph styled as Page Break with manual page break inserted
'        Selection.Find.ClearFormatting
'        With Selection.Find
'            .Text = "^m^p"
'            .Replacement.Text = "^m^p"
'            .Forward = True
'            .Wrap = wdFindStop
'            .Format = True
'            .Style = ActiveDocument.Styles("Page Break (pb)")
'            .MatchCase = False
'            .MatchWholeWord = False
'            .MatchWildcards = False
'            .MatchSoundsLike = False
'            .MatchAllWordForms = False
'        End With
'
'    Do While Selection.Find.Execute = True And kCount < 200            'jCount so we don't get an infinite loop
'        kCount = kCount + 1
'        nCount = 0
'        'select following paragraph
'        Selection.Next(Unit:=wdParagraph, Count:=1).Select
'        nextStyle = Selection.Style
'        pageNumK = Selection.Information(wdActiveEndPageNumber)
'
'           For N = LBound(arrSecStartStyles()) To UBound(arrSecStartStyles())
'                'Check if preceding paragraph style is correct
'                If nextStyle <> arrSecStartStyles(N) Then
'                    nCount = nCount + 1
'                Else
'                    Exit For
'                End If
'            Next N
'
'            If nCount = UBound(arrSecStartStyles()) Then
'                kString = kString & "** ERROR: " & nextStyle & " style on page " & pageNumK _
'                    & " cannot follow Page Break (pb) style." & vbNewLine & vbNewLine
'            End If
'
'        'Debug.Print kString
'
'Err2Resume:
'
'        'move the selection back to original paragraph, so it won't be
'        'selected again on next search
'        Selection.Previous(Unit:=wdParagraph, Count:=1).Select
'    Loop
'
'    'Debug.Print kString
'
'    CheckAfterPB = kString
'
'    Exit Function
'
'ErrHandler1:
'    If Err.Number = 5941 Or Err.Number = 5834 Then       'Style doesn't exist in document
'        Exit Function
'    End If
'
'ErrHandler2:
'    If Err.Number = 5941 Or Err.Number = 5834 Then       ' Style doesn't exist in document
'        Resume Err2Resume
'    End If
'
'End Function
'
'
'Private Function FixTrackChanges() As Boolean
'    Dim N As Long
'    Dim oComments As Comments
'    Set oComments = ActiveDocument.Comments
'
'    'See if there are tracked changes or comments in document
'    On Error Resume Next
'    Selection.HomeKey Unit:=wdStory   'start search at beginning of doc
'    WordBasic.NextChangeOrComment       'search for a tracked change or comment. error if none are found.
'
'    'If there are changes, ask user if they want macro to accept changes or cancel
'    If Err = 0 Then
'        If MsgBox("Bookmaker doesn't like comments or tracked changes, but it appears that you have some in your document." _
'            & vbCr & vbCr & "Click OK to ACCEPT ALL CHANGES and DELETE ALL COMMENTS right now and continue with the Bookmaker Requirements Check." _
'            & vbCr & vbCr & "Click CANCEL to stop the Bookmaker Requirements Check and deal with the tracked changes and comments on your own.", _
'            273, "Are those tracked changes I see?") = vbCancel Then           '273 = vbOkCancel(1) + vbCritical(16) + vbDefaultButton2(256)
'                FixTrackChanges = False
'                Exit Function
'        Else 'User clicked OK, so accept all tracked changes and delete all comments
'            ActiveDocument.AcceptAllRevisions
'            For N = oComments.Count To 1 Step -1
'                oComments(N).Delete
'            Next N
'            Set oComments = Nothing
'        End If
'    End If
'
'End Function
'
'Private Function BadTorStyles(ProgressBar2 As ProgressBar, StatusBar As String, ProgressTitle As String, Stories() As Variant) As String
'    'Called from GoodBadStyles sub if torDOTcom parameter is set to True.
'
'    Dim paraStyle As String
'    Dim activeParaCount As Integer
'
'    Dim strCsvFileName As String
'    Dim strLogInfo() As Variant
'    ReDim strLogInfo(1 To 3)
'    Dim strFullPathToCsv As String
'    Dim arrTorStyles() As Variant
'    Dim strLogDir As String
'    Dim strPathToLogFile As String
'
'    Dim intBadCount As Integer
'    Dim activeParaRange As Range
'    Dim pageNumber As Integer
'
'    Dim N As Integer
'    Dim M As Integer
'    Dim strBadStyles As String
'    Dim A As Long
'
'    Dim TheOS As String
'    TheOS = System.OperatingSystem
'    Dim sglPercentComplete As Single
'    Dim strStatus As String
'
'    Application.ScreenUpdating = False
'
'
'    ' This is the file we want to download
'    strCsvFileName = "Styles_Bookmaker.csv"
'
'    ' We need the info about the log file for any download
'    strLogInfo() = CreateLogFileInfo(FileName:=strCsvFileName)
'    strLogDir = strLogInfo(2)
'    strPathToLogFile = strLogInfo(3)
'    strFullPathToCsv = strLogDir & Application.PathSeparator & strCsvFileName
'
'    ' download the list of good Tor styles from Confluence
'    Dim downloadStyles As GitBranch
'    ' switch to develop for testing
'    downloadStyles = master
'
'    If DownloadFromConfluence(DownloadSource:=downloadStyles, _
'                                FinalDir:=strLogDir, _
'                                LogFile:=strPathToLogFile, _
'                                FileName:=strCsvFileName) = False Then
'        ' If it's False, DL failed. Is a previous version there?
'        If IsItThere(strFullPathToCsv) = False Then
'            ' Sorry can't DL right now, no previous file in directory
'            MsgBox "Sorry, I can't download the Bookmaker style info right now."
'            Exit Function
'        Else
'            ' Can't DL new file but old one exists, let's use that
'            MsgBox "I can't download the Bookmaker style info right now, so I'll just use the old info I have on file."
'        End If
'    End If
'
'
'    'List of styles approved for use in Bookmaker
'    'Organized by approximate frequency in manuscripts (most freq at top)
'    'returned array is dimensioned with 1 column, need to specify row and column (base 0)
'    arrTorStyles = LoadCSVtoArray(Path:=strFullPathToCsv, RemoveHeaderRow:=True, RemoveHeaderCol:=False)
'
'    activeParaCount = ActiveDocument.Paragraphs.Count
'
'    For N = 1 To activeParaCount
'
'
'        If N Mod 100 = 0 Then
'            'Percent complete and status for progress bar (PC) and status bar (Mac)
'            sglPercentComplete = (((N / activeParaCount) * 0.1) + 0.76)
'            strStatus = "* Checking paragraph " & N & " of " & activeParaCount & " for approved Bookmaker styles..." & vbCr & StatusBar
'
'            Call UpdateBarAndWait(Bar:=ProgressBar2, Status:=strStatus, Percent:=sglPercentComplete)
'        End If
'
'        For A = LBound(Stories()) To UBound(Stories())
'            If N <= ActiveDocument.StoryRanges(Stories(A)).Paragraphs.Count Then
'                paraStyle = ActiveDocument.StoryRanges(Stories(A)).Paragraphs(N).Style
'                'Debug.Print paraStyle
'
'                If Right(paraStyle, 1) = ")" Then
'                    'Debug.Print "Current paragraph is: " & paraStyle
'                    On Error GoTo ErrHandler
'
'                    intBadCount = -1        ' -1 because the array is base 0
'
'                    For M = LBound(arrTorStyles()) To UBound(arrTorStyles())
'                        'Debug.Print arrTorStyles(M, 0)
'
'                        If paraStyle <> arrTorStyles(M, 0) Then
'                            intBadCount = intBadCount + 1
'                        Else
'                            Exit For
'                        End If
'                    Next M
'
'                    'Debug.Print intBadCount
'                    If intBadCount = UBound(arrTorStyles()) Then
'                        Set activeParaRange = ActiveDocument.StoryRanges(A).Paragraphs(N).Range
'                        pageNumber = activeParaRange.Information(wdActiveEndPageNumber)
'                        strBadStyles = strBadStyles & "** ERROR: Non-Bookmaker style on page " & pageNumber _
'                            & " (Paragraph " & N & "):  " & paraStyle & vbNewLine & vbNewLine
'                            'Debug.Print strBadStyles
'                    End If
'
'                End If
'            End If
'        Next A
'ErrResume:
'
'    Next N
'
'    StatusBar = "* Checking paragraphs for approved Bookmaker styles..." & vbCr & StatusBar
'
'    'Debug.Print strBadStyles
'
'    BadTorStyles = strBadStyles
'    Exit Function
'
'ErrHandler:
'    Debug.Print Err.Number & " " & Err.Description & " | " & Err.HelpContext
'    If Err.Number = 5941 Or Err.Number = 5834 Then       'style is not in document
'        Resume ErrResume
'    End If
'
'End Function
'
'Private Function CountReqdStyles() As Variant
'    Dim arrStyleName(1 To 15) As String                      ' Declare number of items in array
'    Dim intStyleCount() As Variant
'    ReDim intStyleCount(1 To 15) As Variant                  ' Delcare items in array. Must be dynamic to pass back to Sub
'
'    Dim A As Long
'    Dim xCount As Integer
'
'    Application.ScreenUpdating = False
'
'    arrStyleName(1) = "Titlepage Book Title (tit)"
'    arrStyleName(2) = "Titlepage Author Name (au)"
'    arrStyleName(3) = "span ISBN (isbn)"
'    arrStyleName(4) = "Chap Number (cn)"
'    arrStyleName(5) = "Chap Title (ct)"
'    arrStyleName(6) = "Chap Title Nonprinting (ctnp)"
'    arrStyleName(7) = "Titlepage Logo (logo)"
'    arrStyleName(8) = "Part Title (pt)"
'    arrStyleName(9) = "Part Number (pn)"
'    arrStyleName(10) = "FM Head (fmh)"
'    arrStyleName(11) = "FM Title (fmt)"
'    arrStyleName(12) = "BM Head (bmh)"
'    arrStyleName(13) = "BM Title (bmt)"
'    arrStyleName(14) = "Illustration holder (ill)"
'    arrStyleName(15) = "Illustration Source (is)"
'
'    For A = 1 To UBound(arrStyleName())
'        On Error GoTo ErrHandler
'        intStyleCount(A) = 0
'        With ActiveDocument.Range.Find
'            .ClearFormatting
'            .Text = ""
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindStop
'            .Format = True
'            .Style = ActiveDocument.Styles(arrStyleName(A))
'            .MatchCase = False
'            .MatchWholeWord = False
'            .MatchWildcards = False
'            .MatchSoundsLike = False
'            .MatchAllWordForms = False
'        Do While .Execute(Forward:=True) = True And intStyleCount(A) < 100   ' < 100 to prevent infinite loop, especially if content controls in title or author blocks
'            intStyleCount(A) = intStyleCount(A) + 1
'        Loop
'        End With
'ErrResume:
'    Next
'
'
'    '------------Exit Sub if exactly 100 Titles counted, suggests hidden content controls-----
'    If intStyleCount(1) = 100 Then
'
'        MsgBox "Something went wrong!" & vbCr & vbCr & "It looks like you might have content controls (form fields or drop downs) in your document, but Word for Mac doesn't play nicely with these." _
'        & vbCr & vbCr & "Try running this macro on a PC or contact workflows@macmillan.com for assistance.", vbCritical, "OH NO!!"
'        Exit Function
'
'    End If
'
'    'For A = 1 To UBound(arrStyleName())
'    '    Debug.Print arrStyleName(A) & ": " & intStyleCount(A) & vbNewLine
'    'Next A
'
'    CountReqdStyles = intStyleCount()
'    Exit Function
'
'ErrHandler:
'    If Err.Number = 5941 Or Err.Number = 5834 Then
'        intStyleCount(A) = 0
'        Resume ErrResume
'    End If
'
'End Function
'
'Private Sub FixSectionHeadings(oldStyle As String, newStyle As String)
'
'    Application.ScreenUpdating = False
'
'    'check if styles exist, else exit sub
'    On Error GoTo ErrHandler:
'    Dim keyStyle As Word.Style
'
'    Set keyStyle = ActiveDocument.Styles(oldStyle)
'    Set keyStyle = ActiveDocument.Styles(newStyle)
'
'    'Move selection to start of document
'    Selection.HomeKey Unit:=wdStory
'
'        'Find paras styles as CN and change to CT style
'        Selection.Find.ClearFormatting
'        Selection.Find.Style = ActiveDocument.Styles(oldStyle)
'        Selection.Find.Replacement.ClearFormatting
'        Selection.Find.Replacement.Style = ActiveDocument.Styles(newStyle)
'        With Selection.Find
'            .Text = ""
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = True
'            .MatchCase = False
'            .MatchWholeWord = False
'            .MatchWildcards = False
'            .MatchSoundsLike = False
'            .MatchAllWordForms = False
'        End With
'        Selection.Find.Execute Replace:=wdReplaceAll
'
'Exit Sub
'
'ErrHandler:
'    If Err.Number = 5941 Or Err.Number = 5834 Then 'the requested member of the collection does not exist (i.e., style doesn't exist)
'        Exit Sub
'    End If
'
'End Sub
'
'Private Function GetMetadata() As String
'    Dim styleNameB(3) As String         ' must declare number of items in array here
'    Dim bString(3) As String            ' and here
'    Dim B As Integer
'    Dim strTitleData As String
'
'    Application.ScreenUpdating = False
'
'    styleNameB(1) = "Titlepage Book Title (tit)"
'    styleNameB(2) = "Titlepage Author Name (au)"
'    styleNameB(3) = "span ISBN (isbn)"
'
'    For B = 1 To UBound(styleNameB())
'        bString(B) = GetText(styleNameB(B))
'        If bString(B) <> vbNullString Then
'            bString(B) = "** " & styleNameB(B) & " **" & vbNewLine & _
'                        bString(B) & vbNewLine
'        End If
'
'        strTitleData = strTitleData & bString(B)
'
'    Next B
'
'    'Debug.Print strTitleData
'
'    GetMetadata = strTitleData
'
'End Function
'
'Private Function IllustrationsList() As String
'    Dim cString(1000) As String             'Max number of illustrations. Could be lower than 1000.
'    Dim cCount As Integer
'    Dim pageNumberC As Integer
'    Dim strFullList As String
'    Dim N As Integer
'    Dim strSearchStyle As String
'
'    Application.ScreenUpdating = False
'
'    strSearchStyle = "Illustration holder (ill)"
'    cCount = 0
'
'    'Move selection to start of document
'    Selection.HomeKey Unit:=wdStory
'
'        ' Check if search style exists in document
'        On Error GoTo ErrHandler
'        Dim keyStyle As Style
'
'        Set keyStyle = ActiveDocument.Styles(strSearchStyle)
'
'        Selection.Find.ClearFormatting
'        With Selection.Find
'            .Text = ""
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindStop
'            .Format = True
'            .Style = ActiveDocument.Styles(strSearchStyle)
'            .MatchCase = False
'            .MatchWholeWord = False
'            .MatchWildcards = False
'            .MatchSoundsLike = False
'            .MatchAllWordForms = False
'        End With
'
'    Do While Selection.Find.Execute = True And cCount < 1000            'cCount < 1000 so we don't get an infinite loop
'        cCount = cCount + 1
'        pageNumberC = Selection.Information(wdActiveEndPageNumber)
'
'        'If paragraph return exists in selection, don't select last character (the last paragraph return)
'        If InStr(Selection.Text, Chr(13)) > 0 Then
'            Selection.MoveEnd Unit:=wdCharacter, Count:=-1
'        End If
'
'        cString(cCount) = "Page " & pageNumberC & ": " & Selection.Text & vbNewLine
'
'        'If the next character is a paragraph return, add that to the selection
'        'Otherwise the next Find will just select the same text with the paragraph return
'        Selection.MoveEndWhile Cset:=Chr(13), Count:=wdForward
'
'    Loop
'
'    'Move selection back to start of document
'    Selection.HomeKey Unit:=wdStory
'
'    If cCount > 1000 Then
'        MsgBox "You have more than 1,000 illustrations tagged in your manuscript." & vbNewLine & _
'        "Please contact workflows@macmillan.com to complete your illustration list."
'    End If
'
'    If cCount = 0 Then
'        cCount = 1
'        cString(1) = "no illustrations detected" & vbNewLine
'    End If
'
'    For N = 1 To cCount
'        strFullList = strFullList & cString(N)
'    Next N
'
'    'Debug.Print strFullList
'
'    IllustrationsList = strFullList
'
'    Exit Function
'
'ErrHandler:
'    If Err.Number = 5941 Or Err.Number = 5834 Then
'        IllustrationsList = ""
'        Exit Function
'    End If
'
'End Function
'
'Function CheckPrev2Paras(StyleA As String, StyleB As String, StyleC As String) As String
'    Dim strErrors As String
'    Dim intCount As Integer
'    Dim pageNum As Integer
'    Dim intCurrentPara As Integer
'    Dim strStyle1 As String
'    Dim strStyle2 As String
'    Dim strStyle3 As String
'
'    Application.ScreenUpdating = False
'
'        'check if styles exist, else exit sub
'        On Error GoTo ErrHandler:
'        Dim keyStyle As Word.Style
'
'        Set keyStyle = ActiveDocument.Styles(StyleA)
'        Set keyStyle = ActiveDocument.Styles(StyleB)
'        Set keyStyle = ActiveDocument.Styles(StyleC)
'
'
'    strErrors = ""
'
'    'Move selection to start of document
'    Selection.HomeKey Unit:=wdStory
'
'    'select paragraph with that style
'        Selection.Find.ClearFormatting
'        With Selection.Find
'            .Text = ""
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindStop
'            .Format = True
'            .Style = ActiveDocument.Styles(StyleC)
'            .MatchCase = False
'            .MatchWholeWord = False
'            .MatchWildcards = False
'            .MatchSoundsLike = False
'            .MatchAllWordForms = False
'        End With
'
'    intCount = 0
'
'    Do While Selection.Find.Execute = True And intCount < 300            'jCount < 300 so we don't get an infinite loop
'        intCount = intCount + 1
'
'        'Get number of current pagaraph, because we get an error if try to select before 1st para
'
'        intCurrentPara = ActiveDocument.Range(0, Selection.Paragraphs(1).Range.End).Paragraphs.Count
'
'        'Debug.Print intCurrentPara
'
'        'Also determine if selection is the LAST paragraph of the document, for later
'        Dim SelectionIncludesFinalParagraphMark As Boolean
'        If Selection.Type = wdSelectionNormal And Selection.End = ActiveDocument.Content.End Then
'            SelectionIncludesFinalParagraphMark = True
'        Else
'            SelectionIncludesFinalParagraphMark = False
'        End If
'
'        'Debug.Print intCurrentPara
'
'        If intCurrentPara > 1 Then      'NOT first paragraph of document
'            'select preceding paragraph
'            Selection.Previous(Unit:=wdParagraph, Count:=1).Select
'            pageNum = Selection.Information(wdActiveEndPageNumber)
'
'                'Check if preceding paragraph style is correct
'                If Selection.Style <> StyleA Then
'
'                    If Selection.Style = StyleB Then
'                        'select preceding paragraph again, see if it's prevStyle
'                        Selection.Previous(Unit:=wdParagraph, Count:=1).Select
'                        pageNum = Selection.Information(wdActiveEndPageNumber)
'
'                            If Selection.Style <> StyleA Then
'                                strErrors = strErrors & "** ERROR: " & StyleB & " followed by " & StyleC & "" _
'                                    & " on" & vbNewLine & vbTab & "page " & pageNum & " must be preceded by " _
'                                    & StyleA & "." & vbNewLine & vbNewLine
'                            Else
'                                'If you're searching for a page break before, also check if manual page break is in paragraph
'                                If StyleA = "Page Break (pb)" Then
'                                    If InStr(Selection.Text, Chr(12)) = 0 Then
'                                        strErrors = strErrors & "** ERROR: Missing manual page break on page " & pageNum & "." _
'                                            & vbNewLine & vbNewLine
'                                    End If
'                                End If
'                            End If
'
'                        Selection.Next(Unit:=wdParagraph, Count:=1).Select
'                    Else
'
'                        strErrors = strErrors & "** ERROR: " & StyleC & " on page " _
'                            & pageNum & " must be used after an" & vbNewLine & vbTab & StyleA & "." _
'                                & vbNewLine & vbNewLine
'
'                    End If
'                Else
'                    'Make sure initial selection wasn't last paragraph, or else we'll error when trying to select after it
'                    If SelectionIncludesFinalParagraphMark = False Then
'                        'select follow paragraph again, see if it's a Caption
'                        Selection.Next(Unit:=wdParagraph, Count:=2).Select
'                        pageNum = Selection.Information(wdActiveEndPageNumber)
'
'                            If Selection.Style = StyleB Then
'                                strErrors = strErrors & "** ERROR: " & StyleC & " style on page " & pageNum & " must" _
'                                    & " come after " & StyleB & " style." & vbNewLine & vbNewLine
'                            End If
'                        Selection.Previous(Unit:=wdParagraph, Count:=2).Select
'                    End If
'
'                    'If you're searching for a page break before, also check if manual page break is in paragraph
'                    If StyleA = "Page Break (pb)" Then
'                        If InStr(Selection.Text, Chr(12)) = 0 Then
'                            strErrors = strErrors & "** ERROR: Missing manual page break on page " & pageNum & "." _
'                                & vbNewLine & vbNewLine
'                        End If
'                    End If
'                End If
'
'                'Debug.Print strErrors
'
'            'move the selection back to original paragraph, so it won't be
'            'selected again on next search
'            Selection.Next(Unit:=wdParagraph, Count:=1).Select
'
'        Else 'Selection is first paragraph of the document
'            strErrors = strErrors & "** ERROR: " & StyleC & " cannot be first paragraph of document." & vbNewLine & vbNewLine
'        End If
'
'    Loop
'
'    '------------------------Search for Illustration holder and check previous paragraph--------------
'    'Move selection to start of document
'    Selection.HomeKey Unit:=wdStory
'
'    'select paragraph with that style
'        Selection.Find.ClearFormatting
'        With Selection.Find
'            .Text = ""
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindStop
'            .Format = True
'            .Style = ActiveDocument.Styles(StyleA)
'            .MatchCase = False
'            .MatchWholeWord = False
'            .MatchWildcards = False
'            .MatchSoundsLike = False
'            .MatchAllWordForms = False
'        End With
'
'    intCount = 0
'
'    Do While Selection.Find.Execute = True And intCount < 1000            'jCount < 1000 so we don't get an infinite loop
'        intCount = intCount + 1
'
'        'Get number of current pagaraph, because we get an error if try to select before 1st para
'        intCurrentPara = ActiveDocument.Range(0, Selection.Paragraphs(1).Range.End).Paragraphs.Count
'
'        If intCurrentPara > 1 Then      'NOT first paragraph of document
'            'select preceding paragraph
'            Selection.Previous(Unit:=wdParagraph, Count:=1).Select
'            pageNum = Selection.Information(wdActiveEndPageNumber)
'
'                'Check if preceding paragraph style is a Caption, which is not allowed
'                If Selection.Style = StyleB Then
'                    strErrors = strErrors & "** ERROR: " & StyleB & " on page " & pageNum & " must come after " _
'                                    & StyleA & "." & vbNewLine & vbNewLine
'                End If
'
'            Selection.Next(Unit:=wdParagraph, Count:=1).Select
'        End If
'    Loop
'
'    'Debug.Print strErrors
'
'    CheckPrev2Paras = strErrors
'
'    'Move selection back to start of document
'    Selection.HomeKey Unit:=wdStory
'    Exit Function
'
'ErrHandler:
'    If Err.Number = 5941 Or Err.Number = 5834 Then       'Style doesn't exist in document
'        Exit Function
'    End If
'
'End Function
'
'Private Function CreateReportText(TemplateUsed As Boolean, errorList As String, metadata As String, illustrations As String, goodStyles As String) As String
'
'    Application.ScreenUpdating = False
'
'    Dim strReportText As String
'
'    If TemplateUsed = False Then
'        strReportText = strReportText & vbNewLine
'        strReportText = strReportText & "------------------------STYLES IN USE--------------------------" & vbNewLine
'        strReportText = strReportText & "It looks like you aren't using the Macmillan style template." & vbNewLine
'        strReportText = strReportText & "That's OK, but if you would like more info about your document," & vbNewLine
'        strReportText = strReportText & "just attach the Macmillan style template and apply the styles" & vbNewLine
'        strReportText = strReportText & "throughout the document." & vbNewLine
'        strReportText = strReportText & vbNewLine
'        strReportText = strReportText & goodStyles
'    Else
'        If errorList = "" Then
'            strReportText = strReportText & vbNewLine
'            strReportText = strReportText & "                 CONGRATULATIONS! YOU PASSED!" & vbNewLine
'            strReportText = strReportText & " But you're not done yet. Please check the info listed below." & vbNewLine
'            strReportText = strReportText & vbNewLine
'        Else
'            strReportText = strReportText & vbNewLine
'            strReportText = strReportText & "                             OOPS!" & vbNewLine
'            strReportText = strReportText & "     Problems were found with the styles in your document." & vbNewLine
'            strReportText = strReportText & vbNewLine
'            strReportText = strReportText & vbNewLine
'            strReportText = strReportText & "---------------------------- ERRORS ---------------------------" & vbNewLine
'            strReportText = strReportText & errorList
'            strReportText = strReportText & vbNewLine
'            strReportText = strReportText & vbNewLine
'        End If
'            strReportText = strReportText & "--------------------------- METADATA --------------------------" & vbNewLine
'            strReportText = strReportText & "If any of the information below is wrong, please fix the" & vbNewLine
'            strReportText = strReportText & "associated styles in the manuscript." & vbNewLine
'            strReportText = strReportText & vbNewLine
'            strReportText = strReportText & metadata
'            strReportText = strReportText & vbNewLine
'            strReportText = strReportText & vbNewLine
'            strReportText = strReportText & "----------------------- ILLUSTRATION LIST ---------------------" & vbNewLine
'
'            If illustrations <> "no illustrations detected" & vbNewLine Then
'                strReportText = strReportText & "Verify that this list of illustrations includes only the file" & vbNewLine
'                strReportText = strReportText & "names of your illustrations." & vbNewLine
'                strReportText = strReportText & vbNewLine
'            End If
'
'            strReportText = strReportText & illustrations
'            strReportText = strReportText & vbNewLine
'            strReportText = strReportText & vbNewLine
'            strReportText = strReportText & "-------------------- MACMILLAN STYLES IN USE ------------------" & vbNewLine
'            strReportText = strReportText & goodStyles
'    End If
'
'    CreateReportText = strReportText
'
'End Function
'
'Private Function StylesInUse(ProgressBar As ProgressBar, Status As String, ProgTitle As String, Stories() As Variant) As String
'    'Creates a list of all styles in use, not just Macmillan styles
'    'No list of bad styles
'    'For use when no Macmillan template is attached
'
'    Dim TheOS As String
'    TheOS = System.OperatingSystem
'    Dim sglPercentComplete As Single
'    Dim strStatus As String
'
'    Dim activeDoc As Document
'    Set activeDoc = ActiveDocument
'    Dim stylesGood() As String
'    Dim stylesGoodLong As Long
'    stylesGoodLong = 400                                    'could maybe reduce this number
'    ReDim stylesGood(stylesGoodLong)
'    Dim styleGoodCount As Integer
'    Dim activeParaCount As Integer
'    Dim J As Integer, K As Integer, L As Integer
'    Dim paraStyle As String
'    '''''''''''''''''''''
'    Dim activeParaRange As Range
'    Dim pageNumber As Integer
'    Dim A As Long
'
'    '----------Collect all styles being used-------------------------------
'    styleGoodCount = 0
'    activeParaCount = activeDoc.Paragraphs.Count
'    For J = 1 To activeParaCount
'
'        'All Progress Bar statements for PC only because won't run modeless on Mac
'        If J Mod 100 = 0 Then
'
'            'Percent complete and status for progress bar (PC) and status bar (Mac)
'            sglPercentComplete = (((J / activeParaCount) * 0.12) + 0.86)
'            strStatus = "* Checking paragraph " & J & " of " & activeParaCount & " for Macmillan styles..." & vbCr & Status
'
'            Call UpdateBarAndWait(Bar:=ProgressBar, Status:=strStatus, Percent:=sglPercentComplete)
'
'        End If
'
'        For A = LBound(Stories()) To UBound(Stories())
'            If J <= ActiveDocument.StoryRanges(Stories(A)).Paragraphs.Count Then
'                paraStyle = activeDoc.StoryRanges(Stories(A)).Paragraphs(J).Style
'                Set activeParaRange = activeDoc.StoryRanges(Stories(A)).Paragraphs(J).Range
'                pageNumber = activeParaRange.Information(wdActiveEndPageNumber)                 'alt: (wdActiveEndAdjustedPageNumber)
'
'                For K = 1 To styleGoodCount
'                    ' "Left" function because now stylesGood includes page number, so won't match paraStyle
'                    If paraStyle = Left(stylesGood(K), InStrRev(stylesGood(K), " --") - 1) Then
'                        K = styleGoodCount                              'stylereport bug fix #1    v. 3.1
'                        Exit For                                        'stylereport bug fix #1   v. 3.1
'                    End If                                              'stylereport bug fix #1   v. 3.1
'                Next K
'                If K = styleGoodCount + 1 Then
'                    styleGoodCount = K
'                    stylesGood(styleGoodCount) = paraStyle & " -- p. " & pageNumber
'                End If
'            End If
'        Next A
'    Next J
'
'    'Sort good styles
'    If K <> 0 Then
'    ReDim Preserve stylesGood(1 To styleGoodCount)
'    WordBasic.SortArray stylesGood()
'    End If
'
'    'Create single string for good styles
'    Dim strGoodStyles As String
'    For K = LBound(stylesGood()) To UBound(stylesGood())
'        strGoodStyles = strGoodStyles & stylesGood(K) & vbNewLine
'    Next K
'
'    'Debug.Print strGoodStyles
'
'    StylesInUse = strGoodStyles
'
'End Function
'
Private Sub ISBNcleanup()
'removes "span ISBN (isbn)" style from all but the actual ISBN numerals

    'check if that style exists, if not then exit sub
    On Error GoTo ErrHandler:
        Dim keyStyle As Word.Style
        Set keyStyle = ActiveDocument.Styles("span ISBN (isbn)")

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
  'Style doesn't exist in document
    If Err.Number = 5941 Or Err.Number = 5834 Then
      Exit Sub
    Else
      Err.Source = strReports & "ISBNcleanup"
      If genUtils.GeneralHelpers.ErrorChecker(Err) = False Then
        Resume
      Else
        Call genUtils.Reports.ReportsTerminate
      End If
    End If

End Sub

'Private Function BookTypeCheck()
'    ' Validates the book types listed following the ISBN on the copyright page.
'    Dim intCount As Integer
'    Dim strErrors As String
'    Dim strBookTypes(1 To 7) As String
'    Dim A As Long
'    Dim blnMissing As Boolean
'    Dim strISBN As String
'
'    strBookTypes(1) = "trade paperback"
'    strBookTypes(2) = "hardcover"
'    strBookTypes(3) = "e-book"
'    strBookTypes(4) = "ebook"
'    strBookTypes(5) = "print on demand"
'    strBookTypes(6) = "print-on-demand"
'    strBookTypes(7) = "mass market paperback"
'
'    'Move selection back to start of document
'    Selection.HomeKey Unit:=wdStory
'
'    On Error GoTo ErrHandler
'
'    intCount = 0
'    With Selection.Find
'        .ClearFormatting
'        .Text = ""
'        .Replacement.Text = ""
'        .Forward = True
'        .Wrap = wdFindStop
'        .Format = True
'        .Style = ActiveDocument.Styles("span ISBN (isbn)")
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'
'        Do While .Execute(Forward:=True) = True And intCount < 100   ' < 100 to precent infinite loop
'            intCount = intCount + 1
'            strISBN = Selection.Text
'            'Record current selection because we need to return to it later
'            ActiveDocument.Bookmarks.Add Name:="ISBN", Range:=Selection.Range
'
'            Selection.Collapse Direction:=wdCollapseEnd
'            Selection.EndOf Unit:=wdLine, Extend:=wdExtend
'
'            blnMissing = True
'                For A = 1 To UBound(strBookTypes())
'                    If InStr(Selection.Text, "(" & strBookTypes(A) & ")") > 0 Then
'                        blnMissing = False
'                        Exit For
'                    End If
'                Next A
'
'            If blnMissing = True Then
'                strErrors = strErrors & "** ERROR: Correct book type required in parentheses after" & vbNewLine & _
'                    "ISBN " & strISBN & " on copyright page." _
'                    & vbNewLine & vbNewLine
'            End If
'
'            'Now we need to return the selection to where it was above, or else we can't loop through selection.find
'            If ActiveDocument.Bookmarks.Exists("ISBN") = True Then
'                Selection.GoTo what:=wdGoToBookmark, Name:="ISBN"
'                ActiveDocument.Bookmarks("ISBN").Delete
'            End If
'
'        Loop
'
'    End With
'
'    'Debug.Print strErrors
'    BookTypeCheck = strErrors
'    Exit Function
'
'ErrHandler:
'    Err.Source = strReports & "BookTypeCheck"
'  ' style doesn't exist in document
'    If Err.Number = 5941 Or Err.Number = 5834 Then
'        Exit Function
'    End If
'
'End Function
'
'Private Function CheckNonprintingText()
'    ' Verify that all "Chapter Title Nonprinting (ctnp)" paragraphs have some body text
'    Dim iCount As Long
'    Dim strBodyText As String
'    Dim strErrors As String
'    Dim pageNum As Long
'    Dim intCount As Long
'
'
'    'Move selection back to start of document
'    Selection.HomeKey Unit:=wdStory
'
'    On Error GoTo ErrHandler
'
'    intCount = 0
'    With Selection.Find
'        .ClearFormatting
'        .Text = ""
'        .Replacement.Text = ""
'        .Forward = True
'        .Wrap = wdFindStop
'        .Format = True
'        .Style = ActiveDocument.Styles("Chap Title Nonprinting (ctnp)")
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'
'        Do While .Execute(Forward:=True) = True And intCount < 1000   ' < 1000 to precent infinite loop
'            intCount = intCount + 1
'            strBodyText = Selection.Text
'
'            pageNum = Selection.Information(wdActiveEndPageNumber)
'
''            'Record current selection because we need to return to it later
''            ActiveDocument.Bookmarks.Add Name:="CTNP", Range:=Selection.Range
''
''            Selection.Collapse Direction:=wdCollapseEnd
''            Selection.EndOf Unit:=wdLine, Extend:=wdExtend
'
'            If strBodyText = Chr(13) Then
'                strErrors = strErrors & _
'                    "** ERROR: Chap Title Nonprinting paragraph on page " & pageNum & " requires body text." & _
'                    vbNewLine & vbNewLine
'            End If
'
'
''            'Now we need to return the selection to where it was above, or else we can't loop through selection.find
''            If ActiveDocument.Bookmarks.Exists("ISBN") = True Then
''                Selection.GoTo what:=wdGoToBookmark, Name:="ISBN"
''                ActiveDocument.Bookmarks("ISBN").Delete
''            End If
'
'        Loop
'
'    End With
'
'    CheckNonprintingText = strErrors
'
'    Exit Function
'
'ErrHandler:
'        'Debug.Print Err.Number & ": " & Err.Description
'    If Err.Number = 5941 Or Err.Number = 5834 Then      ' style doesn't exist in document
'        Exit Function
'    End If
'
'End Function
'
'Private Sub ChapNumCleanUp()
'    ' Removes character styles from Chapter Number paragraphs
'    Dim iCount As Long
'    Dim strText As String
'    Dim intCount As Long
'
'    'Move selection back to start of document
'    Selection.HomeKey Unit:=wdStory
'
'    On Error GoTo ErrHandler
'
'    intCount = 0
'    With Selection.Find
'        .ClearFormatting
'        .Text = ""
'        .Replacement.Text = ""
'        .Forward = True
'        .Wrap = wdFindStop
'        .Format = True
'        .Style = ActiveDocument.Styles("Chap Number (cn)")
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'
'        Do While .Execute(Forward:=True) = True And intCount < 1000   ' < 1000 to precent infinite loop
'            intCount = intCount + 1
'            #If Mac Then
'                ' Mac 2011 doesn't support ClearCharacterFormattingAll method
'                ' And ClearFormatting removes paragraph formatting as well
'                Selection.ClearFormatting
'                Selection.Style = "Chap Number (cn)"
'            #Else
'                Selection.ClearCharacterAllFormatting
'            #End If
'        Loop
'
'    End With
'
'
'    Exit Sub
'
'ErrHandler:
'        'Debug.Print Err.Number & ": " & Err.Description
'    If Err.Number = 5941 Or Err.Number = 5834 Then      ' style doesn't exist in document
'        Exit Sub
'    End If
'End Sub
'
'
'Private Function CheckFileName() As Boolean
'' Returns error message if file name contains special characters
'
'    Dim strDocName As String
'    Dim strCheckChar As String
'    Dim strAllGoodChars As String
'    Dim lngNameLength As Long
'    Dim R As Long
'    Dim strErrorString As String
'
'    CheckFileName = False
'
'    ' Only alphanumeric, underscore and hyphen allowed in Bkmkr names
'    ' Will do vbTextCompare later for case insensitive search
'    strAllGoodChars = "ABCDEFGHIJKLMNOPQRSTUVWZYX1234567890_-"
'
'    ' Get file name w/o extension
'    strDocName = ActiveDocument.Name
'    strDocName = Left(strDocName, InStrRev(strDocName, ".") - 1)
'
'    lngNameLength = Len(strDocName)
'
'    ' Loop: pull each char in file name, check if it appears in good char
'    ' list. If it doesn't appear, then it's bad! So return True
'    ' Error is same whether there is 1 or 100 bad chars, so exit as soon as
'    ' one is found.
'
'    For R = 1 To lngNameLength
'        strCheckChar = Mid(strDocName, R, 1)
'        If InStr(1, strAllGoodChars, strCheckChar, vbTextCompare) = 0 Then
'            CheckFileName = True
'            Exit Function
'        End If
'    Next R
'
'End Function
