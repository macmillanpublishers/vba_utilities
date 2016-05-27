Attribute VB_Name = "CastoffMacro"
Option Explicit
Sub UniversalCastoff()
' created by Erica Warren - erica.warren@macmillan.com

' ========== PUROPOSE ========================
' Takes user inputs from userform to calculate a castoff (estimated print page count) based on
' the current text of the document.

' ========== DEPENDENCIES ====================
' 1. Requires SharedMacros module to be installed in same template
' 2. Requires design character count CSV files and spine size files be saved as attachments to
'    https://confluence.macmillan.com/display/PBL/Word+Template+downloads+-+production
'    - info on how to format
' 3. Requires CastoffForm userform module and TextBoxEventHandler class module. Input validation is done in CastoffForm.
    
    ' ======= Check if EN and FN stories =====
    Dim stStories() As Variant
    stStories = StoryArray

    ' ======= Run startup checks ========
    ' True means a check failed (e.g., doc protection on)
    If StartupSettings(StoriesUsed:=stStories) = True Then
        Call Cleanup
        Exit Sub
    End If
    
    '----------Load userform------------------------
    Dim objCastoffForm As CastoffForm
    Set objCastoffForm = New CastoffForm
    
    objCastoffForm.Show
    
End Sub

Public Sub CastoffStart(FormInputs As CastoffForm)
                
    ' Get estimated page count
    Dim lngCastoffResult() As Long
    
    If FormInputs.chkDesignPickup.value = True Then
        ' If it's a pickup, use the PickupDesign math to get the page count, only 1 result
        ReDim Preserve lngCastoffResult(0)
        lngCastoffResult(0) = PickupDesign(FormInputs)
    Else
        ' Get designs selected from Form.
        Dim lngDesign() As Long     ' Index number of design density in CSV file, starts at 0
        Dim strDesign() As String   ' Text of design density
        Dim lngDim As Long       ' Number of dimensions of lngDesign and strDesign, base 1
        
        lngDim = -1     ' so we start base 0
        
        'For each design checked, increase dimension by 1 and then assign index and text of the design to an array
        If FormInputs.chkDesignLoose Then
            lngDim = lngDim + 1
            ReDim Preserve lngDesign(0 To lngDim)
            ReDim Preserve strDesign(0 To lngDim)
            lngDesign(lngDim) = 0
            strDesign(lngDim) = FormInputs.chkDesignLoose.Caption
        End If
        
        If FormInputs.chkDesignAverage Then
            lngDim = lngDim + 1
            ReDim Preserve lngDesign(0 To lngDim)
            ReDim Preserve strDesign(0 To lngDim)
            lngDesign(lngDim) = 1
            strDesign(lngDim) = FormInputs.chkDesignAverage.Caption
        End If
        
        If FormInputs.chkDesignTight Then
            lngDim = lngDim + 1
            ReDim Preserve lngDesign(0 To lngDim)
            ReDim Preserve strDesign(0 To lngDim)
            lngDesign(lngDim) = 2
            strDesign(lngDim) = FormInputs.chkDesignTight.Caption
        End If

        
        '---------Download CSV with design specs from Confluence site-------
        Dim arrDesign() As Variant
        Dim strCastoffFile As String    'File name of CSV on Confluence
    
        'Create name of castoff csv file to download
        strCastoffFile = "Castoff_" & FormInputs.PublisherCode & ".csv"
        
        arrDesign = DownloadCSV(FileName:=strCastoffFile)
        
        ' Check that returned array is allocated
        If IsArrayEmpty(arrDesign) = True Then
            Exit Sub ' Error messages were in DownloadCSV (and DownloadFromConfluence) so none needed here
        End If
        
        '------------Get castoff for each Design selected-------------------
        Dim D As Long
        
        For D = LBound(lngDesign()) To UBound(lngDesign())
            
            'Error handling: lngDesign(d) must be in range of design array
            If UBound(arrDesign(), 1) <= lngDesign(D) And UBound(arrDesign(), 2) <= FormInputs.TrimIndex Then
                 MsgBox "There was an error generating your castoff. Please contact workflows@macmillan.com for assistance.", _
                    vbCritical, "Error 1: Design Count Out of Range"
                Unload FormInputs
                Exit Sub
            Else
    
                '---------Calculate Page Count--------------------------------------
                ReDim Preserve lngCastoffResult(D)
                lngCastoffResult(D) = Castoff(lngDesign(D), arrDesign(), FormInputs)
                
            End If
        Next D
        
        ' ----- Get spine size if POD -------

'        Dim strSpineSize As String
'        strSpineSize = ""
'
'        If FormInputs.PrintType = FormInputs.optPrintPOD.Caption Then
'            strSpineSize = SpineSize(FormInputs.Staging, lngCastoffResult(0))
'            'Debug.Print "spine size = " & strSpineSize
'        End If

    End If
    
    '-------------Create final message---------------------------------------------------
    Dim strReportText As String
        
    ' Create text of castoff from arrays
    Dim strCastoffs As String
    Dim strPickupTitle As String
    Dim E As Long
    
    ' If it's a pickup, there is only 1 option
    If FormInputs.chkDesignPickup Then
        strCastoffs = lngCastoffResult(0)
        strPickupTitle = "PICKUP TITLE: " & FormInputs.txtPrevTitle_pickup.value & vbNewLine
    Else
        strPickupTitle = ""
        strCastoffs = vbNewLine
        For E = LBound(lngCastoffResult) To UBound(lngCastoffResult)
            strCastoffs = strCastoffs & vbTab & strDesign(E) & ": " & lngCastoffResult(E) & vbNewLine
        Next E
    End If
    
    
    strReportText = _
    " * * * MACMILLAN PRELIMINARY CASTOFF * * * " & vbNewLine & _
    vbNewLine & _
    "DATE: " & Date & vbNewLine & _
    "TITLE: " & FormInputs.txtTitle & vbNewLine & _
    "AUTHOR: " & FormInputs.txtAuthor & vbNewLine & _
    "PUBLISHER: " & FormInputs.Imprint & vbNewLine & _
    "EDITOR: " & FormInputs.txtEditor & vbNewLine & _
    "TRIM SIZE: " & FormInputs.TrimSize & vbNewLine & _
    vbNewLine & _
    strPickupTitle & _
    "SCHEDULED PAGE COUNT: " & FormInputs.numTxtPageCount & vbNewLine & _
    "ESTIMATED PAGE COUNT: " & _
    strCastoffs
    
    '-------------Report castoff info to user----------------------------------------------------------------
    Call CreateTextFile(strText:=strReportText, suffix:="Castoff")
    
    Call Cleanup
    Unload FormInputs
            
End Sub

Private Function Castoff(lngDesignIndex As Long, arrCSV() As Variant, objForm As CastoffForm) As Long
    
    ' Get total CHARACTER count (incl. notes and spaces) from document
    Dim lngTotalCharCount As Long
    lngTotalCharCount = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharactersWithSpaces, _
                        IncludeFootnotesAndEndnotes:=True)
                        
    ' Get CHARACTER count for text without embedded endnotes and footnotes
    Dim lngNotesCharCount As Long
    lngNotesCharCount = lngTotalCharCount - ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharactersWithSpaces, _
                        IncludeFootnotesAndEndnotes:=False)
          
    ' Get PAGE count of main text
    Dim lngMainTextPages As Long
    ActiveDocument.Repaginate
    lngMainTextPages = ActiveDocument.StoryRanges(wdMainTextStory).Information(wdNumberOfPagesInDocument)
        
    ' Get PAGE count of endnotes and footnotes
    Dim lngEndnotesPages As Long
    Dim lngFootnotesPages As Long
    Dim lngNotesPages As Long
    
    If EndnotesExist = True Then
        ActiveDocument.Repaginate
        lngEndnotesPages = ActiveDocument.StoryRanges(wdEndnotesStory).Information(wdNumberOfPagesInDocument)
    Else
        lngEndnotesPages = 0
    End If
    
    If FootnotesExist = True Then
        ActiveDocument.Repaginate
        lngFootnotesPages = ActiveDocument.StoryRanges(wdFootnotesStory).Information(wdNumberOfPagesInDocument)
    Else
        lngFootnotesPages = 0
    End If
    
    lngNotesPages = lngEndnotesPages + lngFootnotesPages
                    
    ' Calculate number of characters per page of NOTES in the MANUSCRIPT if there are linked notes
    ' If there aren't linked notes, get char/page of whole manuscript (because we can't divide by 0)
    Dim lngMsCharPerPage As Long
    If lngNotesPages > 0 Then
        lngMsCharPerPage = lngNotesCharCount / lngNotesPages
    Else
        lngMsCharPerPage = lngTotalCharCount / lngMainTextPages
    End If
    
    ' Get number of unlinked endnotes pages in MS from Form, estimate number of characters
    Dim lngUnlinkedNotesCharCount As Long
    ' Form code validates that empty string = 0
    lngUnlinkedNotesCharCount = objForm.numTxtUnlinkedNotes_std * lngMsCharPerPage

    
    ' Get number of endnotes TK from Form, estimate number of characters
    Dim lngEndnotesTKCharCount As Long
    ' Form code validates that empty string = 0
    lngEndnotesTKCharCount = objForm.numTxtNotesTK_std * lngMsCharPerPage
    
    ' Get number of biblio pages in manuscript from Form, estimate number of characters
    Dim lngBiblioMsCharCount As Long
    ' Form code validates that empty string = 0
    lngBiblioMsCharCount = objForm.numTxtBibliography_std * lngMsCharPerPage
    
    ' Get number of biblio pages TK from Form, estimate number of characters
    Dim lngBiblioTKCharCount As Long
    ' Form code validates that empty string = 0
    lngBiblioTKCharCount = objForm.numTxtBiblioTK_std * lngMsCharPerPage
    
    ' Calculate total character count of main text and notes separately
    Dim lngMainCharCount As Long
    lngMainCharCount = lngTotalCharCount - lngNotesCharCount - lngUnlinkedNotesCharCount - lngBiblioMsCharCount
    lngNotesCharCount = lngNotesCharCount + lngUnlinkedNotesCharCount + lngBiblioMsCharCount + lngEndnotesTKCharCount _
                        + lngBiblioTKCharCount
                            
    ' --------------------------------------------------
    ' For Reference: Index numbers in arrCSV (base 0)
    '
    '         | 5-1/2 x 8-1/4 |  6-1/8 x 9-1/4
    'loose    | (0,0)         | (0,1)
    'average  | (1,0)         | (1,1)
    'tight    | (2,0)         | (2,1)
    'notes    | (3,0)         | (3,1)
    'lines    | (4,0)         | (4,1)
    'overflow | (5,0)         | (5,1)
    '--------------------------------------------------
    
    '---------Get design character count from CSV-------------------------------
    Dim lngDesignCount As Long
    lngDesignCount = arrCSV(lngDesignIndex, objForm.TrimIndex)
    'Debug.Print lngDesignCount
    
    '---------Get notes character count from CSV--------------------------------
    Dim lngNotesDesign As Long
    lngNotesDesign = arrCSV(3, objForm.TrimIndex)   ' notes always in position 3
    
    '---------Get lines per page from CSV--------------------------------------
    Dim lngLinesPage As Long
    lngLinesPage = arrCSV(4, objForm.TrimIndex)     ' lines per page always in position 4
    
    '---------Get overflow pages from CSV--------------------------------------
    Dim lngOverflow As Long
    lngOverflow = arrCSV(5, objForm.TrimIndex)      ' overflow pages always in position 5

    '----------Get user inputs from Userform--------------------------------------------------
    
    ' The rest of the inputs are not required, but are assigned 0 if left empty in form
    
    ' Calculate number of pages!
    Dim lngMainPages As Long
    Dim lngTotalNotesPages As Long
    Dim lngPartsPages As Long
    Dim lngHeadingPages As Long
    Dim lngEstPages As Long
    
    lngMainPages = lngMainCharCount / lngDesignCount
    lngTotalNotesPages = lngNotesCharCount / lngNotesDesign
    lngPartsPages = objForm.numTxtParts_std * 2
    ' 3 (below) because headings take up 3 lines each
    ' 2 because we ask for headings in 2 chapters
    lngHeadingPages = ((objForm.numTxtSubheads_std / 2) * objForm.numTxtChapters_std * 3) / lngLinesPage
    
    lngEstPages = lngMainPages _
                + lngTotalNotesPages _
                + lngPartsPages _
                + lngHeadingPages _
                + objForm.numTxtChapters_std _
                + objForm.numTxtFrontmatter_std _
                + objForm.numTxtIndex_std _
                + objForm.numTxtBackmatter_std _
                + objForm.numTxtTables_std _
                + objForm.numTxtArt_std
                
    Dim lngFinalResult As Long
    
    lngFinalResult = FinalSig(lngEstPages, objForm)
    
    Castoff = lngFinalResult

End Function

Private Function SpineSize(PageCount As Long)
    ' right now, for POD titles only
    ' which we're not even tracking anymore, but leaving code here just in case
    
    Dim strSpine As String
    
    If PageCount < 48 Then
        strSpine = "NOTE: POD titles less than 48 pages will be saddle-stitched."
    ElseIf PageCount >= 48 And PageCount <= 1050 Then       'Limits of spine size table
    
        '----Define spine chart file name--------------------------------
        Dim strSpineFile As String
        'strSpineFile = "Spine_" & Publisher & ".csv"   ' Use this if we're doing different paper based on imprint
        strSpineFile = "POD_Spines.csv"                 ' Use this if we're doing 1 kind of paper for all POD
        
        Dim arrSpine() As Variant
        
        arrSpine = DownloadCSV(FileName:=strSpineFile)
        
        ' Check that returned array is allocated
        If IsArrayEmpty(arrSpine) = True Then
            strSpine = "ERROR: cannot calculate spine size"
            Exit Function ' Error messages were in DownloadCSV (and DownloadFromConfluence) so none needed here
        End If
    
        '---------Lookup spine size in array-------------------------------
        Dim C As Long
        
        For C = LBound(arrSpine, 1) To UBound(arrSpine, 1)
            'Debug.Print arrDesign(c, 0) & " = " & PageCount
            If arrSpine(C, 0) = PageCount Then
                strSpine = arrSpine(C, 1)
                Exit For
            End If
        Next C
    Else
        strSpine = "Your page count of " & PageCount & _
                " is out of range of the spine-size table."
    End If
    
    
    If strSpine = vbNullString Then
        strSpine = "Word was unable to generate a spine size. " & _
            "Contact workflows@macmillan.com for assistance."
    Else
        strSpine = "Your spine size will be " & strSpine & " inches " & _
                            "at this page count."
    End If

    'Debug.Print strSpine
    SpineSize = strSpine

End Function

Private Function PickupDesign(objCastoffForm As CastoffForm) As Long
' estimate page count based on design of previous book
    
    ' get total character count of current ms from document
    ' this includes EN/FN, but at some point could calculate those differently like we do in Castoff
    Dim lngCurrentMsCharCount As Long
    lngCurrentMsCharCount = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharactersWithSpaces, _
                        IncludeFootnotesAndEndnotes:=True)
    
    ' divide total prev character count by page count to get avg characters per page in prev book
    Dim lngPrevCharPerBookPage As Long
    lngPrevCharPerBookPage = objCastoffForm.numTxtPrevCharCount_pickup.value / objCastoffForm.numTxtPrevPageCount_pickup.value
    
    ' divide total characters of this doc by avg characters per page to get est page count, add additional pages
    Dim lngStartingResult As Long
    lngStartingResult = (lngCurrentMsCharCount / lngPrevCharPerBookPage) + objCastoffForm.numTxtAddlPgs_pickup.value
    
    ' Calculate what the final sig will be
    Dim lngFinalPageCount As Long
    lngFinalPageCount = FinalSig(lngStartingResult, objCastoffForm)
    
    PickupDesign = lngFinalPageCount
    
End Function

    
Private Function FinalSig(RawEstPages As Long, objCastForm As CastoffForm) As Long
    ' Figure out what the final sig/page count will be
    Dim result As Long
           
'    If objCastForm.PrintType = objCastForm.optPrintPOD.Caption Then
'        'POD only has to be even, not 16-page sig
'        If (RawEstPages Mod 2) = 0 Then      'page count is even
'            result = RawEstPages
'        Else                                    'page count is odd
'            result = RawEstPages + 1
'        End If
'    Else 'It's printing offset, already validated in castoff form code
        
        ' Calculate next sig up and next sig down
        Dim lngRemainderPgs As Long
        Dim lngLowerSig As Long
        Dim lngUpperSig As Long
        
        lngRemainderPgs = RawEstPages Mod 16
        lngLowerSig = RawEstPages - lngRemainderPgs
        lngUpperSig = RawEstPages + (16 - lngRemainderPgs)
                    
        ' Get number of overflow pages from CSV
        ' DL again even if we just did, because if it's a pickup we didn't DL
        ' later improvement: have it check date of last DL and only DL once a day
        Dim arrCastoff() As Variant
        Dim strFile As String
        
        strFile = "Castoff_" & objCastForm.PublisherCode & ".csv"
        
        arrCastoff = DownloadCSV(FileName:=strFile)
        
        Dim lngOverflow As Long
        lngOverflow = arrCastoff(5, objCastForm.TrimIndex)    ' 5 is index of overflow info in CSV
        
        ' Determine if we go up or down a signature
        If lngRemainderPgs < lngOverflow Then
            result = lngLowerSig
        Else
            result = lngUpperSig
        End If
'    End If

    FinalSig = result
    
End Function


Private Function DownloadCSV(FileName As String, Optional DownloadFrom As GitBranch = master) As Variant
    '---------Download CSV with design specs from Confluence site-------

    'Create log file name
    Dim arrLogInfo() As Variant
    ReDim arrLogInfo(1 To 3)
    
    arrLogInfo() = CreateLogFileInfo(FileName)
      
    'Create final path for downloaded CSV file (in log directory)
    'not in temp dir because that is where DownloadFromConfluence downloads it to, and it cleans that file up when done
    Dim strStyleDir As String
    Dim strPath As String
    Dim strLogFile As String
    Dim strMessage As String
    Dim strDir As String
    
    strStyleDir = arrLogInfo(1)
    strDir = arrLogInfo(2)
    strLogFile = arrLogInfo(3)
    strPath = strDir & Application.PathSeparator & FileName
        
    'Check if log file already exists; if not, create it
    CheckLog strStyleDir, strDir, strLogFile
    
    'Download CSV file from Confluence
    If DownloadFromConfluence(FinalDir:=strDir, LogFile:=strLogFile, FileName:=FileName, DownloadSource:=DownloadFrom) = False Then
        ' If download fails, check if we have an older version of the CSV to work with
        If IsItThere(strPath) = False Then
            strMessage = "Looks like we can't download the design info from the internet right now. " & _
                "Please check your internet connection, or contact workflows@macmillan.com."
            MsgBox strMessage, vbCritical, "Error 5: Download failed, no previous design file available"
            Exit Function
        Else
            strMessage = "Looks like we can't download the most up-to-date design info from the internet right now, " & _
                "so we'll just use the info we have on file for your castoff."
            MsgBox strMessage, vbInformation, "Let's do this thing!"
        End If
    End If
    
    ' Heading row/col different based on different InfoTypes
    Dim blnRemoveHeaderRow As Boolean
    Dim blnRemoveHeaderCol As Boolean
    
    ' Because the castoff CSV has header row and col, but Spine CSV only has a header row
    If InStr(1, FileName, "Castoff") <> 0 Then
        blnRemoveHeaderRow = True
        blnRemoveHeaderCol = True
    ElseIf InStr(1, FileName, "Spine") <> 0 Then
        blnRemoveHeaderRow = True
        blnRemoveHeaderCol = False
    End If
    
    'Double check that CSV is there
    Dim arrFinal() As Variant
    If IsItThere(strPath) = False Then
        strMessage = "The Castoff macro is unable to access the design count file right now. Please check your internet " & _
                    "connection and try again, or contact workflows@macmillan.com."
        MsgBox strMessage, vbCritical, "Error 3: Design CSV doesn't exist"
        Exit Function
    Else
        ' Load CSV into an array
        arrFinal = LoadCSVtoArray(Path:=strPath, RemoveHeaderRow:=blnRemoveHeaderRow, RemoveHeaderCol:=blnRemoveHeaderCol)
    End If
    
    DownloadCSV = arrFinal
    
End Function
