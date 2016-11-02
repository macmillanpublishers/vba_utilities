Attribute VB_Name = "Endnotes"
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       ENDNOTES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ====== PURPOSE ==============================================================
' Unlink embedded endnotes and add to new section in main doc, with correct
' headings and numbering

' ====== DEPENDENCIES ============
' 1. Manuscript must be styled with Macmillan custom styles.
' 2. Assumes section breaks have already been added at end of each section.


' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    DECLARATIONS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Option Base 1

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    GLOBAL VARIABLES and CONSTANTS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Const c_strEndnotes As String = "genUtils.Endnotes."

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    PUBLIC PROCEDURES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ===== EndnoteCheck ==========================================================
' Call this function to run automated endnote cleanup for validator.

Public Function EndnoteCheck() As genUtils.Dictionary
  On Error GoTo EndnoteCheckError
  
  Dim dictReturn As genUtils.Dictionary
  Set dictReturn = New genUtils.Dictionary
  
  Dim blnNotesExist As Boolean
  blnNotesExist = NotesExist()
  dictReturn.Add "endnotesExist", blnNotesExist
  
  If blnNotesExist = True Then
    Set dictReturn = EndnoteUnlink(p_blnAutomated:=True)
  Else
    dictReturn.Add "pass", True
  End If

  Set EndnoteCheck = dictReturn
  Exit Function

EndnoteCheckError:
  Err.Source = c_strEndnotes & "EndnoteCheck"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.Reports.ReportsTerminate
  End If
End Function

' ===== EndnoteDeEmbed ========================================================
' Call this sub if being run by a person (by clicking macro button), not
' automatically on server. Can't combine this and EndnoteCheck because that
' needs to be a function, this needs to be a sub.

Public Sub EndnoteDeEmbed()
  Set activeDoc = activeDoc

  Dim blnNotesExist As Boolean
  blnNotesExist = NotesExist()
  
  If blnNotesExist = True Then
    Dim dictStep As genUtils.Dictionary
    Set dictStep = EndnoteUnlink(p_blnAutomated:=False)
  Else
    MsgBox "Sorry, no linked endnotes found in document. Click OK to exit" _
      & " the Endnotes macro."
  End If
  
  ' Eventually do something with the dictionary (log?)

End Sub


' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    PRIVATE PROCEDURES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ===== NotesExist ============================================================
' Are there even endnotes?

Private Function NotesExist() As Boolean
  If activeDoc.Endnotes.Count > 0 Then
    NotesExist = True
  Else
    NotesExist = False
  End If
End Function

' ===== EndnoteUnlink =========================================================
' Unlinks embedded endnotes and places them in their own section at the end of
' the document, with headings for each chapter. Note numbers restart at 1 for
' each chapter.

Private Function EndnoteUnlink(p_blnAutomated As Boolean) As genUtils.Dictionary
  On Error GoTo EndnoteUnlinkError
  
  ' --------- Declare and set variables ---------------
  Dim dictReturn As genUtils.Dictionary
  Set dictReturn = New genUtils.Dictionary
  dictReturn.Add "pass", False

  Dim palgraveTag As Boolean
  Dim iReply As Integer
  Dim sglPercentComplete As Single
  Dim strStatus As String
  Dim strTitle As String
  palgraveTag = False

  '-----------Turn off track changes--------
  Dim currentTracking As Boolean
  currentTracking = activeDoc.TrackRevisions
  activeDoc.TrackRevisions = False
  
' -----------------------------------------------------------------------------
' -----------------------------------------------------------------------------
' This section only if being run by a person.
' -----------------------------------------------------------------------------
' -----------------------------------------------------------------------------
  If p_blnAutomated = False Then

  ' ------ Doesn't work on Mac ---------------
  ' TEST THIS NOW THAT IT'S BEEN REFACTORED...
    #If Mac Then
      MsgBox "It looks like you are on a Mac. Unfortunately, this macro only" & _
        " works properly on Windows. Click OK to exit the Endnotes macro."
      Exit Function
    #End If
    
    If activeDoc.Sections.Count = 1 Then
      iReply = MsgBox("Only one section found in document. Without section " & _
        "breaks, endnotes will be numbered continuously from beginning to end." _
        & vbNewLine & vbNewLine & "If you would like to continue without " & _
        "section breaks, click OK." & vbNewLine & "If you would like to exit " & _
        "the macro and add section breaks at the end of each chapter to " & _
        "trigger note numbering to restart at 1 for each chapter, click Cancel.", _
        vbYesNo + vbExclamation + vbDefaultButton2, "Alert")
      If iReply = vbNo Then
        Exit Function
      End If
    End If
    
    ' ----- See if we're using custom Palgrave tags -----
    iReply = MsgBox("To insert bracketed <NoteCallout> tags around your " & _
      "endnote references, click YES." & vbNewLine & vbNewLine & "To " & _
      "continue with standard superscripted endnote reference numbers only," & _
      " click NO.", vbYesNo + vbExclamation + vbDefaultButton2, "Alert")
    If iReply = vbYes Then palgraveTag = True

    '------------record status of current status bar and then turn on-------
    Dim currentStatusBar As Boolean
    currentStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
  
    '--------Progress Bar------------------------------
    'Percent complete and status for progress bar (PC) and status bar (Mac)
    'Requires ProgressBar custom UserForm and Class
    strTitle = "Unlink Endnotes"
    sglPercentComplete = 0.04
    strStatus = "* Getting started..."
    
    Dim objProgressNotes As ProgressBar
    Set objProgressNotes = New ProgressBar
    
    objProgressNotes.Title = strTitle
    Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=objProgressNotes, _
      Status:=strStatus, Percent:=sglPercentComplete)
  End If
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' END SECTION FOR NON-VALIDATOR VERSION
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------

  ' Begin working on Endnotes
  Application.ScreenUpdating = False
  
  Dim lngTotalSections As Long
  Dim lngTotalNotes As Long
  Dim objSection As Section ' each section obj we're looking through
  Dim objEndnote As Endnote ' each Endnote obj in section
  Dim strFirstStyle As String
  Dim strSecondStyle As String
  Dim rngHeading As Range
  Dim lngNoteNumber As Long ' Integer for the superscripted note number
  Dim lngNoteCount As Long ' count of TOTAL notes in doc
  Dim rngNoteNumber As Range
  Dim strCountMsg As String
  Dim lngSectionCount As Long
  Dim blnAddText As Boolean
  Dim N As Long
  Dim A As Long
  
  lngTotalSections = activeDoc.Sections.Count
  lngTotalNotes = activeDoc.Endnotes.Count
  lngNoteNumber = 1
  lngNoteCount = 0
  lngSectionCount = 0
  
  dictReturn.Add "palgraveTags", palgraveTag
  dictReturn.Add "numSections", lngTotalSections
  dictReturn.Add "numNotes", lngTotalNotes

' ----- ADD NOTES SECTION HEADING TO END OF DOC -------------------------------
  Dim rngNotes As Range
  Set rngNotes = activeDoc.StoryRanges(wdMainTextStory).Paragraphs.Last.Range
  Dim a_strText(1 To 2) As String
  Dim a_strStyle(1 To 2) As String
  
  ' Add page break before new section
  a_strText(1) = vbNewLine
  a_strStyle(1) = Reports.strPageBreak
  
  ' N.B., If your Range includes the final paragraph return of the doc,
  ' wdCollapseEnd leaves the insertion point BEFORE that last paragraph
  ' return character. Hence newline BEFORE text below.
  a_strText(2) = vbNewLine & "Notes"
  a_strStyle(2) = Reports.strBmHead
  
  For A = LBound(a_strText) To UBound(a_strText)
    With rngNotes
      .InsertAfter a_strText(A)
      .Collapse wdCollapseEnd
      .Style = a_strStyle(A)
    End With
  Next A

' ----- Loop through sections -------------------------------------------------
  For Each objSection In activeDoc.Sections
    lngSectionCount = lngSectionCount + 1

  ' If no notes in this section, skip to next
    If objSection.Range.Endnotes.Count > 0 Then
      With objSection.Range
      ' Need to check 1st para style for heading text
        strFirstStyle = .Paragraphs(1).Range.ParagraphStyle

        ' If first paragraph is NOT an approved heading, just continue with notes
        ' and numbering as if it is the same section as previous.
        If Reports.IsHeading(strFirstStyle) = True Then
        ' New section, so restart note numbers at 1
          lngNoteNumber = 1
          Set rngHeading = .Paragraphs(1).Range

          ' If it's a CN / CT combo, get CT as well
          ' THIS ISN'T WORKING YET ...
          If strFirstStyle = Reports.strChapNumber Then
            If .Paragraphs.Count > 1 Then
              strSecondStyle = .Paragraphs(2).Range.ParagraphStyle
              If strSecondStyle = Reports.strChapTitle Then
                rngHeading.Expand Unit:=wdParagraph
              End If
            End If
          End If

        ' Add that text as a subhead to final notes section
          blnAddText = AddNoteText(p_rngNoteBody:=rngHeading, p_blnHeading:=True)
          dictReturn.Add objSection.Index & "_NoteHeadAdded", blnAddText
        End If
      End With
      
    ' ----- Now loop through all notes in this section ------------------------
    ' ----- and add to Notes section ------------------------------------------
    ' reset N from last section
      N = 1
      
      For N = 1 To objSection.Range.Endnotes.Count
        lngNoteCount = lngNoteCount + 1
        
      ' ----- Update progress bar if run by user ------------------------------
        If p_blnAutomated = False Then
          If lngNoteCount Mod 10 = 0 Then
            sglPercentComplete = (((lngNoteCount / lngTotalNotes) * 0.95) + 0.04)
            strCountMsg = "* Unlinking endnote " & lngNoteCount & " of " & _
              lngTotalNotes & vbNewLine & strStatus
            Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=objProgressNotes, _
              Status:=strCountMsg, Percent:=sglPercentComplete)
          End If
        End If
      
      ' ----- Add note text to Notes section ----------------------------------
      ' Endnote.Range returns the Range object of the end-of-book note
        Set objEndnote = objSection.Range.Endnotes(N)
      ' Add note text to end Notes section
        Call AddNoteText(p_rngNoteBody:=objEndnote.Range, _
          p_lngNoteNumber:=lngNoteNumber)
      
      ' ----- Add note number to text -----------------------------------------
      ' Endnote.Reference returns Range object of the in-text note number.
        Set rngNoteNumber = objEndnote.Reference
        If palgraveTag = False Then
          rngNoteNumber.InsertAfter lngNoteNumber
        Else
          rngNoteNumber.InsertAfter "<NoteCallout>" & lngNoteNumber & "</NoteCallout>"
        End If
        rngNoteNumber.Style = "span superscript characters (sup)"
      
      ' Increment note number counter
        lngNoteNumber = lngNoteNumber + 1
      Next N
    
    ' ---- Delete notes -------------------------------------------------------
    ' Separate loop than above so we don't mess up our counting in prev. section
      For Each objEndnote In objSection.Range.Endnotes
        objEndnote.Delete
      Next
    End If
  Next objSection
  
' ----- Delete section breaks -------------------------------------------------
' Since we don't need them for endnotes any more. Before rolling this out to
' users though, need to deal with other uses for section breaks (like page numbers)
  genUtils.zz_clearFind
  With activeDoc.Range.Find
    .Text = "^b"  ' Section break character
    .Replacement.Text = vbNullString
  End With

' ---- Test if successful -----------------------------------------------------
  dictReturn.Item("pass") = Not NotesExist()
  
  Set EndnoteUnlink = dictReturn
  
  
  
  activeDoc.TrackRevisions = currentTracking
  Application.DisplayStatusBar = currentStatusBar
  Application.ScreenUpdating = True
  Application.ScreenRefresh
  Exit Function

EndnoteUnlinkError:
  Err.Source = c_strEndnotes & "EndnoteUnlink"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.Reports.ReportsTerminate
  End If
End Function


' ===== AddNoteText ===========================================================
' Adds passed range to Notes section at back of manuscript. Returns if it was
' successful or not. p_lngNoteNumber and p_blnHeading are both optional, but
' must supply one or the other for it to work correctly.

Private Function AddNoteText(p_rngNoteBody As Range, Optional p_lngNoteNumber _
  As Long, Optional p_blnHeading As Boolean = False) As Boolean
  On Error GoTo AddNoteTextError

  Dim objParagraph As Paragraph
  Dim rngNotes As Range
  Dim rngParaText As Range
  Dim strStyle As String
  Dim blnFirstChar As Boolean
  Dim blnLastChar As Boolean
  Dim lngStartChar As Long
  Dim lngEndChar As Long
  Dim strLastChar As String
  Dim B As Long

' ----- Add note number to start of note text ---------------------------------
' N.B., for note text (not headings), p_rngNoteBody is Endnote.Range, which is
' the Range of the end of note text NOT including the embedded note number. But
' p_rngNoteBody.Paragraph(1) DOES include that character (we fix that later).
  If p_blnHeading = False Then
    p_rngNoteBody.InsertBefore p_lngNoteNumber & ". "
  End If
  
' ----- Loop through paragraphs to add each individually ----------------------
' This way we can maintain original paragraph styles.
' For Each objParagraph in p_rngNoteBody.Paragraphs always looped through all
' notes in document, not just paragraphs in this note.
  For B = 1 To p_rngNoteBody.Paragraphs.Count
  ' Set working range at end of main document
    Set rngNotes = activeDoc.StoryRanges(wdMainTextStory).Paragraphs.Last.Range
    Set objParagraph = p_rngNoteBody.Paragraphs(B)
    
  ' Embedded note number is Chr(2); we need to remove thise from the text we're
  ' adding to the main document b/c it appears as a square.
    If objParagraph.Range.Characters.First = Chr(2) Then
      blnFirstChar = True
    Else
      blnFirstChar = False
    End If
  
  ' Chr(13) is paragraph return; we can't insert anything AFTER the final para
  ' return in the document, so we need to remove from insert text to avoid extra
  ' paragraphs. We'll add back if not last paragraph in this note.
    If objParagraph.Range.Characters.Last = Chr(13) Then
      blnLastChar = True
    Else
      blnLastChar = False
    End If
  
  ' If this is a heading and it is MULTIPLE paragraphs, we want to concatenate
  ' into a single line, separated by colon (i.e., CN and CT). Otherwise, we need
  ' to replace the newline removed above.
    If p_blnHeading = True Then
      strLastChar = ": "
      strStyle = "Note Level-1 Subhead (n1)"
    Else
      strLastChar = Chr(13)
      strStyle = objParagraph.Range.ParagraphStyle
    ' Built-in "Endnote Text" style will get converted to "Text - Standard" in
    ' Char Styles macro later if it doesn't have Macmillan code.
      If strStyle = "Endnote Text" Then
        strStyle = strStyle & " (ntx)"
      End If
    End If
  
  ' Calculate where to start and end text to add. Boolean True converts to -1
  ' so we need to take absolute value here. blnFirstChar = True means that we
  ' do NOT want the first character, and thus Abs(blnFirstChar) = 1 and we will
  ' start our new Range at character 2. Ditto for end character.
    lngStartChar = 1 + Abs(blnFirstChar)
    lngEndChar = objParagraph.Range.Characters.Count - Abs(blnLastChar)
    Set rngParaText = objParagraph.Range
    rngParaText.SetRange Start:=rngParaText.Characters(lngStartChar).Start, _
      End:=rngParaText.Characters(lngEndChar).End
  
  ' Copy/paste range text instead of .InsertAfter so we can maintain formatting
    rngParaText.Copy
  
  ' Range right now is final paragraph; reminder we can't add anything AFTER final
  ' paragraph return, so we enter a newline FIRST, then collapse so insertion point
  ' is just before the final paragraph return.
    With rngNotes
      .InsertAfter vbNewLine
      .Collapse wdCollapseEnd
  
    ' Add new paragraph text (with formatting!)
      .PasteAndFormat Type:=wdFormatOriginalFormatting
      .Style = strStyle
    
      If B < p_rngNoteBody.Paragraphs.Count Then
        .InsertAfter strLastChar
      End If
    End With
    
    Set rngNotes = Nothing
    Set objParagraph = Nothing
  Next B

  Exit Function

AddNoteTextError:
  Err.Source = c_strEndnotes & "AddNoteText"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.Reports.ReportsTerminate
  End If
End Function
