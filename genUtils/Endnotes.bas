Attribute VB_Name = "Endnotes"
Option Explicit
Option Base 1
Dim activeRng As Range

Sub EndnoteDeEmbed()

'    '------- Check if document is saved ---------
'    If CheckSave = True Then
'        Exit Sub
'    End If
    
    ' --------- Declare variables ---------------
    Dim refRng As Range
    Dim refSection As Integer
    Dim lastRefSection As Integer
    Dim chapterName As String
    Dim addChapterName As Boolean
    Dim addHeader As Boolean
    Dim nRng As Range, eNote As Endnote, nref As String, refCopy As String
    Dim sectionCount As Long
    Dim StoryRange As Range
    Dim EndnotesExist As Boolean
    Dim TheOS As String
    Dim palgraveTag As Boolean
    Dim iReply As Integer
    Dim BookmarkNum As Integer
    Dim BookmarkName As String
    Dim strCurrentStyle As String
    
    BookmarkNum = 1
    lastRefSection = 0
    addHeader = True
    EndnotesExist = False
    TheOS = System.OperatingSystem
    palgraveTag = False
    
    '''Error checks, setup Doc with sections & numbering
    If TheOS Like "*Mac*" Then
'        MsgBox "It looks like you are on a Mac. Unfortunately, this macro only works properly on Windows. " & _
        "Click OK to exit the Endnotes macro."
        Exit Sub
    End If
    For Each StoryRange In ActiveDocument.StoryRanges
        If StoryRange.StoryType = wdEndnotesStory Then
            EndnotesExist = True
            Exit For
        End If
    Next StoryRange
    If EndnotesExist = False Then
'        MsgBox "Sorry, no linked endnotes found in document. Click OK to exit the Endnotes macro."
        Exit Sub
    End If
    sectionCount = ActiveDocument.Sections.Count
    
    If sectionCount = 1 Then
'        iReply = MsgBox("Only one section found in document. Without section breaks, endnotes will be numbered " & _
        "continuously from beginning to end." & vbNewLine & vbNewLine & "If you would like to continue " & _
        "without section breaks, click OK." & vbNewLine & "If you would like to exit the macro and add " & _
        "section breaks at the end of each chapter to trigger note numbering to restart at 1 for each chapter, click Cancel.", _
        vbYesNo + vbExclamation + vbDefaultButton2, "Alert")
        
'        If iReply = vbNo Then
'            Exit Sub
'        End If
    End If

    '------------record status of current status bar and then turn on-------
    Dim currentStatusBar As Boolean
    currentStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    
    '-----------Turn off track changes--------
    Dim currentTracking As Boolean
    currentTracking = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False
    
    '--------Progress Bar------------------------------
    'Percent complete and status for progress bar (PC) and status bar (Mac)
    'Requires ProgressBar custom UserForm and Class
    Dim sglPercentComplete As Single
    Dim strStatus As String
    Dim strTitle As String
    
    strTitle = "Unlink Endnotes"
    sglPercentComplete = 0.04
    strStatus = "* Getting started..."
    
    #If Mac Then
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    #Else
        Dim objProgressNotes As ProgressBar
        Set objProgressNotes = New ProgressBar
        
        objProgressNotes.Title = strTitle
        objProgressNotes.Show
        
        objProgressNotes.Increment sglPercentComplete, strStatus
        Doze 50 ' Wait 50 milliseconds for progress bar to update
    #End If

    ' Setup global Endnote settings (continuous number, endnotes at document end, number with integers)
    'ActiveDocument.Endnotes.StartingNumber = 1
    'ActiveDocument.Endnotes.NumberingRule = wdRestartContinuous
    ActiveDocument.Endnotes.NumberingRule = wdRestartSection
    ActiveDocument.Endnotes.Location = 1
    ActiveDocument.Endnotes.NumberStyle = wdNoteNumberStyleArabic
    
    
    ' See if we're using custom Palgrave tags
'    iReply = MsgBox("To insert bracketed <NoteCallout> tags around your endnote references, click YES." & vbNewLine & vbNewLine & _
'        "To continue with standard superscripted endnote reference numbers only, click NO.", vbYesNo + vbExclamation + vbDefaultButton2, "Alert")
'    If iReply = vbYes Then palgraveTag = True
    palgraveTag = False
    
    ' Begin working on Endnotes
    Application.ScreenUpdating = False
    
    Dim intNotesCount As Integer
    Dim intCurrentNote As Integer
    Dim strCountMsg As String
    intNotesCount = ActiveDocument.Endnotes.Count
    intCurrentNote = 0
    
    With ActiveDocument
      For Each eNote In .Endnotes
        ' ----- Update progress bar -------------
        intCurrentNote = intCurrentNote + 1
        
        If intCurrentNote Mod 10 = 0 Then
          sglPercentComplete = (((intCurrentNote / intNotesCount) * 0.95) + 0.04)
          strCountMsg = "* Unlinking endnote " & intCurrentNote & " of " & intNotesCount & vbNewLine & strStatus
          
          #If Mac Then
            Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strCountMsg
            DoEvents
          #Else
            objProgressNotes.Increment sglPercentComplete, strCountMsg
            Doze 50 ' Wait 50 milliseconds for progress bar to update
          #End If
        End If
              
        With eNote
          With .Reference.Characters.First
            .Collapse wdCollapseStart
            BookmarkName = "Endnote" & BookmarkNum
            .Bookmarks.Add Name:=BookmarkName
            .InsertCrossReference wdRefTypeEndnote, wdEndnoteNumberFormatted, eNote.Index
            nref = .Characters.First.Fields(1).result
            If palgraveTag = False Then
                .Characters.First.Fields(1).Unlink
            Else
                eNote.Reference.InsertBefore "<NoteCallout>" & nref & "</NoteCallout>"   'tags location of old ref
                .Characters.Last.Fields(1).Delete      ' delete old ref
            End If
    
            'Now for the header business:
            addChapterName = False
            Set refRng = ActiveDocument.Bookmarks(BookmarkName).Range
            refSection = ActiveDocument.Range(0, refRng.Sections(1).Range.End).Sections.Count
            If refSection <> lastRefSection Then
                'following line for debug: comment later
                'MsgBox refSection & " is section of Endnote index #" & nref
                chapterName = endnoteHeader(refSection)
                If chapterName = "```No Header found```" Then
'                    MsgBox "ERROR: Found endnote reference in a section without an approved header style (fmh, cn, ct or ctnp)." & vbNewLine & vbNewLine & _
'                    "Exiting macro, reverting to last save.", vbCritical, "Oh no!"
'                    Documents.Open FileName:=ActiveDocument.FullName, Revert:=True
'                    Application.ScreenUpdating = True
                    Exit Sub
                End If
                addChapterName = True
                lastRefSection = refSection
                'following line for debug: comment later
                'MsgBox chapterName
            End If
            BookmarkNum = BookmarkNum + 1
          End With
          'strCurrentStyle = .Range.Style 'this is to apply save style as orig. note but breaks if more than 1 style.
          .Range.Cut
        End With
        
    '''''Since I am not attempting to number at end of each secion,  commenting out parts of this clause
        'If .Range.EndnoteOptions.Location = wdEndOfSection Then
        '  Set nRng = eNote.Range.Sections.First.Range
        'Else
        Set nRng = .Range
        'End If
        With nRng
          .Collapse wdCollapseEnd
          .End = .End - 1
          If .Characters.Last <> Chr(12) Then .InsertAfter vbCr
          If addHeader = True Then
            .InsertAfter "Notes" & vbCr
            With .Paragraphs.Last.Range
                .Style = "BM Head (bmh)"
            End With
            addHeader = False
          End If
          If addChapterName = True Then
            .InsertAfter chapterName '
            With .Paragraphs.Last.Range
                .Style = "BM Subhead (bmsh)"
            End With
          End If
          .InsertAfter nref & ". "
          With .Paragraphs.Last.Range
            '.Style = strCurrentStyle 'This applies the same style as orig. note, but breaks if more than 1 style used.
            .Style = "Endnote Text"
            .Words.First.Style = "Default Paragraph Font"
          End With
          .Collapse wdCollapseEnd
          .Paste
          If .Characters.Last = Chr(12) Then .InsertAfter vbCr
        End With
      Next
      
      strStatus = "* Unlinking " & intNotesCount & " endnotes..." & vbNewLine & strStatus
      
    '''This deletes the endnote
      For Each eNote In .Endnotes
        eNote.Delete
      Next
    End With
    Set nRng = Nothing
    
    ' ---- apply superscript style to in-text note references -------
    Call zz_clearFind
    Selection.HomeKey wdStory
    
    With Selection.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = ""
      .Replacement.Text = ""
      .Forward = True
      .Wrap = wdFindStop
      .Format = True
      .Style = "Endnote Reference"
      .Replacement.Style = "span superscript characters (sup)"
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .Execute Replace:=wdReplaceAll
    End With
    
        ' ----- Update progress bar -------------
    sglPercentComplete = 0.99
    strStatus = "* Finishing up..." & vbNewLine & strStatus
    
    #If Mac Then
      Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
      DoEvents
    #Else
      objProgressNotes.Increment sglPercentComplete, strStatus
      Doze 50 ' Wait 50 milliseconds for progress bar to update
    #End If
    
    Call RemoveAllBookmarks

Cleanup:
    ' ---- Close progress bar -----
    #If Mac Then
      ' Nothing?
    #Else
      Unload objProgressNotes
    #End If
    
    ActiveDocument.TrackRevisions = currentTracking
    Application.DisplayStatusBar = currentStatusBar
    Application.ScreenUpdating = True
    Application.ScreenRefresh

End Sub

Function endnoteHeader(refSection As Integer) As String
Dim sectionRng As Range
    Dim searchStylesArray(4) As String                       ' number of items in array should be declared here
    Dim searchTest As Boolean
    Dim i As Long
    
    Call zz_clearFind
    
    Set sectionRng = ActiveDocument.Sections(refSection).Range
    searchStylesArray(1) = "FM Head (fmh)"
    searchStylesArray(2) = "Chap Number (cn)"
    searchStylesArray(3) = "Chap Title (ct)"
    searchStylesArray(4) = "Chap Title Nonprinting (ctnp)"
    searchTest = False
    i = 1
    
    Do Until searchTest = True
    Set sectionRng = ActiveDocument.Sections(refSection).Range
    With sectionRng.Find
      .ClearFormatting
      .Style = searchStylesArray(i)
      .Wrap = wdFindStop
      .Forward = True
    End With
    If sectionRng.Find.Execute Then
        endnoteHeader = sectionRng
        searchTest = True
    Else
    'following line for debug: comment later
        'MsgBox searchStylesArray(i) + " Not Found"
        i = i + 1
        If i = 5 Then
            searchTest = True
            endnoteHeader = "```No Header found```"
        End If
    End If
    Loop
        
    Call zz_clearFind
    
End Function

Sub RemoveAllBookmarks()

'three options from http://wordribbon.tips.net/T009004_Removing_All_Bookmarks.html
    'Version 1
    Dim objBookmark As Bookmark
    For Each objBookmark In ActiveDocument.Bookmarks
        objBookmark.Delete
    Next
    
    'Version 2
    'Dim stBookmark As Bookmark
    'ActiveDocument.Bookmarks.ShowHidden = True
    'If ActiveDocument.Bookmarks.Count >= 1 Then
    '   For Each stBookmark In ActiveDocument.Bookmarks
    '      stBookmark.Delete
    '   Next stBookmark
    'End If
    
    'Version 3
    'Dim objBookmark As Bookmark
    '
    'For Each objBookmark In ActiveDocument.Bookmarks
    '    If Left(objBookmark.Name, 1) <> "_" Then objBookmark.Delete
    'Next
    
    
    'http://wordribbon.tips.net/T009004_Removing_All_Bookmarks.html

End Sub


