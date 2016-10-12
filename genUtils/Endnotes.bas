Attribute VB_Name = "Endnotes"
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       ENDNOTES
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ====== PURPOSE ==============================================================
' Manage endnote formatting, primarily for embedded notes.

' ====== DEPENDENCIES ============
' 1. Manuscript must be styled with Macmillan custom styles.


' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    DECLARATIONS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Option Base 1

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    GLOBAL VARIABLES and CONSTANTS
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Const c_strEndnotes As String = "genUtils.Endnotes."
Private g_rngNotes As Range

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
' automatically on server. Can't combine the and EndnoteCheck because that
' needs to be a function, this needs to be a sub.

Public Sub EndnoteDeEmbed()
  Set activeDoc = activeDoc

  Dim blnNotesExist As Boolean
  blnNotesExist = NotesExist()
  
  If blnNotesExist = True Then
    Dim dictStep As genUtils.Dictionary
    Set dictStep = EndnoteUnlink(p_blnAutomated:=False)
  Else
    MsgBox "Sorry, no linked endnotes found in document. Click OK to exit the Endnotes macro."
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
  
' ----------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------
' This section only if being run by a person.
' ----------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------
  If p_blnAutomated = False Then

  ' ------ Doesn't work on Mac ---------------
    #If Mac Then
      MsgBox "It looks like you are on a Mac. Unfortunately, this macro only works properly on Windows. " & _
      "Click OK to exit the Endnotes macro."
      Exit Function
    #End If
    
    If activeDoc.Sections.Count = 1 Then
      iReply = MsgBox("Only one section found in document. Without section breaks, endnotes will be numbered " & _
      "continuously from beginning to end." & vbNewLine & vbNewLine & "If you would like to continue " & _
      "without section breaks, click OK." & vbNewLine & "If you would like to exit the macro and add " & _
      "section breaks at the end of each chapter to trigger note numbering to restart at 1 for each chapter, click Cancel.", _
      vbYesNo + vbExclamation + vbDefaultButton2, "Alert")
      
      If iReply = vbNo Then
          Exit Function
      End If
    End If
    
    ' ----- See if we're using custom Palgrave tags -----
    iReply = MsgBox("To insert bracketed <NoteCallout> tags around your endnote references, click YES." & vbNewLine & vbNewLine & _
        "To continue with standard superscripted endnote reference numbers only, click NO.", vbYesNo + vbExclamation + vbDefaultButton2, "Alert")
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
  
  lngTotalSections = activeDoc.Sections.Count
  lngTotalNotes = activeDoc.Endnotes.Count
  lngNoteNumber = 1
  lngNoteCount = 0
  lngSectionCount = 0
  
  dictReturn.Add "palgraveTags", palgraveTag
  dictReturn.Add "numSections", lngTotalSections
  dictReturn.Add "numNotes", lngTotalNotes
  
' ----- Loop through sections -------------------------------------------------
  For Each objSection In activeDoc.Sections
    lngSectionCount = lngSectionCount + 1
    DebugPrint "Section " & lngSectionCount & " of " & lngTotalSections

  ' If no notes in this section, skip to next
    If objSection.Range.Endnotes.Count > 0 Then
      With objSection.Range
      ' Need to check 1st para style for heading text
        strFirstStyle = .Paragraphs(1).Range.ParagraphStyle
        DebugPrint "First para style: " & strFirstStyle
        ' If first paragraph is not an approved heading, just continue with notes
        ' and numbering as if it is the same section as previous.
        If Reports.IsHeading(strFirstStyle) = True Then
          DebugPrint "Heading!"
        ' New section, so restart note numbers at 1
          lngNoteNumber = 1
          Set rngHeading = .Paragraphs(1).Range
          DebugPrint "Heading text: " & rngHeading.Paragraphs.First.Range.Text
          ' If it's a CN / CT combo, get CT as well
          If strFirstStyle = Reports.strChapNumber Then
            If .Paragraphs.Count > 1 Then
              strSecondStyle = .Paragraphs(2).Range.ParagraphStyle
              If strSecondStyle = Reports.strChapTitle Then
                rngHeading.Expand Unit:=wdParagraph
              End If
            End If
          End If
          DebugPrint "Note heading paragraphs: " & rngHeading.Paragraphs.Count
        ' Add that text as a subhead to final notes section
          blnAddText = AddNoteText(p_rngNoteBody:=rngHeading, p_blnHeading:=True)
          dictReturn.Add "Section" & objSection.Index & "_NoteHeadAdded", _
            blnAddText
        End If
      End With
      
    ' Now loop through all notes in this section and add to Notes section
      DebugPrint objSection.Range.Endnotes.Count
      For Each objEndnote In objSection.Range.Endnotes
      ' ----- Update progress bar if run by user ------------------------------
        lngNoteCount = lngNoteCount + 1
        DebugPrint "Note " & lngNoteCount & " of " & lngTotalNotes
        If p_blnAutomated = False Then
          If lngNoteCount Mod 10 = 0 Then
            sglPercentComplete = (((lngNoteCount / lngTotalNotes) * 0.95) + 0.04)
            strCountMsg = "* Unlinking endnote " & lngNoteCount & " of " & _
              lngTotalNotes & vbNewLine & strStatus
            Call genUtils.ClassHelpers.UpdateBarAndWait(Bar:=objProgressNotes, _
              Status:=strCountMsg, Percent:=sglPercentComplete)
          End If
        End If
    
      ' Add note text to end Notes section
        Call AddNoteText(objEndnote.Range, lngNoteNumber)
      
      ' Add note number to text with superscript style
        Set rngNoteNumber = objEndnote.Reference  ' returns Range of in-text note number
        If palgraveTag = False Then
          rngNoteNumber.InsertAfter lngNoteNumber
        Else
          rngNoteNumber.InsertAfter "<NoteCallout>" & lngNoteNumber & "</NoteCallout>"
        End If
        rngNoteNumber.Style = "span superscript characters (sup)"
      
      ' Increment note number counter
        lngNoteNumber = lngNoteNumber + 1
      Next objEndnote
    
    ' ---- Delete notes in separate loop ----
      For Each objEndnote In objSection.Range.Endnotes
        objEndnote.Delete
      Next
    End If
  Next objSection
  
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
' must supply one or the other.

Private Function AddNoteText(p_rngNoteBody As Range, Optional p_lngNoteNumber _
  As Long = 1, Optional p_blnHeading As Boolean = False) As Boolean
  On Error GoTo AddNoteTextError
  If g_rngNotes Is Nothing Then
' ----- Set up range to hold Notes section we're adding -----------------------
    Dim a_strText(1 To 2) As String
    Dim a_strStyle(1 To 2) As String
    Dim A As Long
    
    a_strText(1) = vbNewLine
    a_strStyle(1) = Reports.strPageBreak
    
    a_strText(2) = "Notes"
    a_strStyle(2) = Reports.strBmHead
    Set g_rngNotes = activeDoc.StoryRanges(wdMainTextStory)
    
    For A = LBound(a_strText) To UBound(a_strText)
      Call AddNewParagraph(p_strText:=a_strText(A), p_strStyle:=a_strStyle(A))
    Next A
  End If

' ----- Add text to that paragraph --------------------------------------------
  Dim objParagraph As Paragraph
  Dim strText As String
  Dim strStyle As String

  If p_blnHeading = False Then
    strText = p_lngNoteNumber & ". "
  End If
  ' Loop through paragraphs to add each individually
    For Each objParagraph In p_rngNoteBody.Paragraphs
      If p_blnHeading = True Then
        strStyle = "Note Level-1 Subhead (n1)"
      Else
        strStyle = objParagraph.Range.ParagraphStyle
      End If
       
      strText = strText & objParagraph.Range.Text
      ' If last char is newline, remove it
      If Right(strText, 1) = Chr(13) Then
        strText = Left(strText, Len(strText) - 1)
      End If
      
      Call AddNewParagraph(p_strText:=strText, p_strStyle:=strStyle)
      strText = vbNullString
    Next objParagraph

  Exit Function

AddNoteTextError:
  Err.Source = c_strEndnotes & "AddNoteText"
  If ErrorChecker(Err, strStyle) = False Then
    Resume
  Else
    Call genUtils.Reports.ReportsTerminate
  End If
End Function

' ===== AddNewParagraph =======================================================
' Add passed text as new paragraph at end of range with the passed style applied
' to the whole paragraph If no range object is passed as argument, then will add
' to global variable. Note that if range is end of DOCUMENT, the last paragraph
' character ALWAYS remains AFTER the collapses range insertion point. This sub
' adds the final newline, so remove from string (though worst case Char Styles
' macro would remove it anyway).

' All objects are passed by ref, so shouldn't need to return the revised range.

Private Sub AddNewParagraph(p_strText As String, p_strStyle As String, _
  Optional p_rngAppend As Range)
  On Error GoTo AddNewParagraphError
  Dim rngNewPara As Range
  Dim blnFinalPara As Boolean
  
  If p_rngAppend Is Nothing Then
    If g_rngNotes Is Nothing Then
      Set g_rngNotes = activeDoc.Paragraphs.Last.Range
    End If
    Set rngNewPara = g_rngNotes
  Else
    Set rngNewPara = p_rngAppend
  End If
  
  With rngNewPara
    rngNewPara.Select
    If GeneralHelpers.ParaIndex() = activeDoc.Paragraphs.Count Then
      blnFinalPara = True
    Else
      blnFinalPara = False
    End If
    
  ' If last para need to add an additional newline, to separate our range from
  ' the previous para
    If blnFinalPara = True Then
      .InsertAfter vbNewLine
      .Collapse wdCollapseEnd
    End If
  
  ' adds new line to keep as a separate paragraph, then move back up to not be
  ' in the new last paragraph, then apply style
    .InsertAfter p_strText & vbNewLine
    .Collapse wdCollapseEnd
    .Move Unit:=wdCharacter, Count:=-1
    .Style = p_strStyle
  End With
  
' Now move original range to end of NEW paragraph
  If p_rngAppend Is Nothing Then
    g_rngNotes.Move Unit:=wdParagraph, Count:=1
    g_rngNotes.Collapse wdCollapseEnd
  Else
    p_rngAppend.Move Unit:=wdParagraph, Count:=1
    p_rngAppend.Collapse wdCollapseEnd
  End If
  Exit Sub

AddNewParagraphError:
  Err.Source = c_strEndnotes & "AddNewParagraph"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.Reports.ReportsTerminate
  End If
End Sub
