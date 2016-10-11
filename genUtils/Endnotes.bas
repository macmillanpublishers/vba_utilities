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
  Dim rngPara As Range
  Dim objEndnote As Endnote ' each Endnote obj in section
  Dim strFirstStyle As String
  Dim strHeading As String
  Dim lngNoteNumber As Long ' Integer for the superscripted note number
  Dim lngNoteCount As Long ' count of TOTAL notes in doc
  Dim rngNoteNumber As Range
  Dim strCountMsg As String
  
  lngTotalSections = activeDoc.Sections.Count
  lngTotalNotes = activeDoc.Endnotes.Count
  lngNoteNumber = 1
  lngNoteCount = 0
  
  dictReturn.Add "palgraveTags", palgraveTag
  dictReturn.Add "numSections", lngTotalSections
  dictReturn.Add "numNotes", lngTotalNotes
  
' ----- Loop through sections -------------------------------------------------
  For Each objSection In activeDoc.Sections
  ' If no notes in this section, skip to next
    If objSection.Range.Endnotes.Count > 0 Then
    ' Need to check 1st para style for heading text, create heading in Notes section
      Set rngPara = objSection.Range.Paragraphs.First.Range
      strFirstStyle = rngPara.ParagraphStyle
      ' If first paragraph is not an approved heading, just continue with notes
      ' and numbering as if it is the same section as previous.
      If Reports.IsHeading(strFirstStyle) = True Then
      ' New section, so restart note numbers at 1
        lngNoteNumber = 1
        strHeading = rngPara.Text
        ' If it's a CN / CT combo, get CT as well
        If strFirstStyle = Reports.strChapNumber Then
          rngPara.Move Unit:=wdParagraph, Count:=1
          If rngPara.ParagraphStyle = Reports.strChapTitle Then
            strHeading = strHeading & ": " & rngPara.Text
          End If
        End If
        strHeading = strHeading & vbNewLine
        ' collapse first, so we can apply style to just-inserted text
        g_rngNotes.Collapse Direction:=wdCollapseEnd
        g_rngNotes.InsertAfter strHeading
        g_rngNotes.Style = "Note Level-1 Subhead (n1)"
      End If
      
    ' Now loop through all notes in this section and add to Notes section
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
' successful or not.

Private Function AddNoteText(p_rngNoteBody As Range, p_lngNoteNumber As Long) _
  As Boolean
  On Error GoTo AddNoteTextError
  If g_rngNotes Is Nothing Then
' ----- Set up range to hold Notes section we're adding -----------------------
    Set g_rngNotes = activeDoc.Range
    g_rngNotes.Collapse wdCollapseEnd
    g_rngNotes.InsertAfter "Notes" & vbNewLine
    g_rngNotes.Style = Reports.strBmHead  ' public constant from Reports module
  End If

  Dim objParagraph As Paragraph
  With g_rngNotes
  ' Collapse range so we can add style after we insert text
    .Collapse Direction:=wdCollapseEnd
    .InsertAfter p_lngNoteNumber & ". "
  ' Loop through paragraphs to add each individually
    For Each objParagraph In p_rngNoteBody.Paragraphs
      .InsertAfter objParagraph.Range.Text
      .Style = objParagraph.Range.ParagraphStyle
      .Collapse Direction:=wdCollapseEnd
    Next objParagraph
  End With
  Exit Function

AddNoteTextError:
  Err.Source = c_strEndnotes & "AddNoteText"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call genUtils.Reports.ReportsTerminate
  End If
End Function


