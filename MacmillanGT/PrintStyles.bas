Attribute VB_Name = "PrintStyles"
Option Explicit
Option Base 1
Sub PrintStyles()
    ' Created by Erica Warren -- erica.warren@macmillan.com
    
    ' places current styles names to the left of each paragraph in the margin and prints document
    
    '=================================================
    '                  Timer Start                  '|
    'Dim StartTime As Double                         '|
    'Dim SecondsElapsed As Double                    '|
                                                    '|
    'Remember time when macro starts                '|
    'StartTime = Timer                               '|
    '=================================================
    
    #If Mac Then
        Call PrintStylesMac
    #Else
        Call PrintStylesPC
    #End If
    
    '============================================================================
    '                   Timer End
    ''''Determine how many seconds code took to run
    'SecondsElapsed = Round(Timer - StartTime, 2)
    
    ''''Notify user in seconds
    'Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
    '============================================================================
    
End Sub

Private Sub PrintStylesMac()

    ' ===== DEPENDENCIES ==========================================================
    ' Requires modules AttachTemplateMacro (and thus also Macmillan style templates)
    ' and ProgressBar userform/class

    
    ' ===== LIMITATIONS ==========================================================
    ' "Normal" can't be in-use as a style in the document
    ' If total margin size (left + right) is < 2 " paragraphs will reflow
    ' Doesn't work for endnotes/footnotes (can't add a drawing object to EN/FNs)
    ' Doesn't work on tables -- breaks the whole macro
    
    ' ====== Check if doc is saved/protected ================
    If CheckSave = True Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' ===== Status Bar ========================
    Dim currentStatusBar As Boolean
    currentStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    
    Dim sglPercentComplete As Single
    Dim strStatus As String
    Dim strTitle As String
    
    strTitle = "Print Styles in Margin"
    sglPercentComplete = 0.01
    strStatus = "* Getting started..."
    
    Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
    DoEvents
    
    ' ===== Copy and Paste into a new doc ===================
    ' Paste below throws an alert that too many styles are being copied, and the option we want is NOT the default
    ' User will have to select "No" from the alert to print correctly
    Dim lngOpt As Long
    
    With Application
    ' record current settings to reset in Cleanup
        lngOpt = .DisplayAlerts
        'lngPasteStyled = .Options.PasteFormatBetweenStyledDocuments ' not available on Mac
        'lngPasteFormat = .Options.PasteFormatBetweenDocuments  ' not available on Mac
        .DisplayAlerts = wdAlertsAll
        '.Options.PasteFormatBetweenStyledDocuments = wdKeepSourceFormatting
        '.Options.PasteFormatBetweenDocuments = wdKeepSourceFormatting
    End With
    
    ' ===== Create new version of this document to manipulate ============
    sglPercentComplete = 0.03
    strStatus = "* Creating dupe document to tag with style names..." & vbNewLine & strStatus
    
    Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
    DoEvents
    
    ' Copy the text of the document into a new document, so we don't screw up the original
    ' Needs to have the BoundMS template attached before copying so the styles match
    ' the new document later, or won't copy any styles
    Dim currentTemplate As String
    Dim CurrentDoc As Document
    Set CurrentDoc = ActiveDocument
    ' Record current template
    currentTemplate = CurrentDoc.AttachedTemplate
    
    ' Attach BoundMS template to original doc, then copy contents
    Call AttachTemplateMacro.zz_AttachBoundMSTemplate
    CurrentDoc.StoryRanges(wdMainTextStory).Copy
    
    Dim tempDoc As Document
    ' Create a new document
    Set tempDoc = Documents.Add(Visible:=False)  ' Visible:=False doesn't work for Mac?
    ' Add Macmillan styles with no color guides 'Might not need this, as we get prompted there are too many styles anyway
    tempDoc.Activate
    Call AttachTemplateMacro.zz_AttachBoundMSTemplate
    'SendKeys not working on Mac for some reason. Need to send "n" (code 45) to the Paste line because pasting a document with
    'many styles triggers a warning and the default behavior is to paste as Normal style. Delay 1 second to allow the message to pop up first.
     MacScript ("delay 1" & vbCr & "tell application " & Chr(34) & "System Events" & Chr(34) & " to key code 45")
     tempDoc.Content.PasteSpecial DataType:=wdPasteHTML 'maintains styles
     


    ' ===== Set margins =================
    sglPercentComplete = 0.05
    strStatus = "* Adjusting margins to fit style names..." & vbNewLine & strStatus
    
    Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
    DoEvents
    
    ' if possible, we want the total margin size to stay the same
    ' so that the paragraphs don't reflow
    ' note margins are in points, 72 pts = 1 inch
    Dim currentLeft As Long
    Dim currentRight As Long
    Dim currentTotal As Long
    
    With tempDoc.PageSetup
        currentLeft = .LeftMargin
        currentRight = .RightMargin
        currentTotal = currentLeft + currentRight
        .LeftMargin = 108   ' 1.5 inches
            If currentTotal >= 144 Then     ' 2 inches
                .RightMargin = currentTotal - 108   ' 1.5 inches
            Else
                .RightMargin = 36   '.5 inches (minimum right margin)
            End If
    End With
    
    ' ==== Change Normal style formatting (since it will define the text boxes) =====
    sglPercentComplete = 0.07
    strStatus = "* Setting format for style names..." & vbNewLine & strStatus

    Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
    DoEvents
        
    ' But save settings first and then change back -- are these settings sticky?
    Dim currentSize As Long
    Dim currentName As String
    Dim currentSpace As Long
    
    With tempDoc.Styles("Normal")
        currentSize = .Font.Size
        currentName = .Font.Name
        currentSpace = .ParagraphFormat.SpaceAfter
        .Font.Size = 7
        .Font.Name = "Calibri"
        .ParagraphFormat.SpaceAfter = 0
    End With
    
    ' ==== Add style names in margin ==================================
    Dim strStyle As String
    Dim lngTextBoxes As Long
    Dim activeParas As Long
    Dim A As Long
    
    ' Count the number of current text boxes etc., because the index number of the new ones
    ' will be offset by that amount
    Dim strStatusLoop As String
    lngTextBoxes = tempDoc.Shapes.Count
    activeParas = tempDoc.Paragraphs.Count
    
    For A = 1 To activeParas
        If A Mod 50 = 0 Then
            sglPercentComplete = Round((((A / activeParas) * 0.85) + 0.1), 2)
            strStatusLoop = "* Adding style names to paragraph " & A & " of " & activeParas & "..." & vbNewLine & strStatus
            
            Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatusLoop
            DoEvents
         
            'SecondsElapsed = Round(Timer - StartTime, 2)
            'Debug.Print "Paragraph " & a & " in " & SecondsElapsed & " seconds"
        End If
    
        tempDoc.Paragraphs(A).Range.Select
        strStyle = Selection.Style
        
        ' Do not tag Text Std to speed up the macro (most common style)
        If strStyle <> "Text - Standard (tx)" Then
            Selection.Collapse Direction:=wdCollapseStart
            Dim newBox As Shape
            Set newBox = tempDoc.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=InchesToPoints(0.13), Top:=0, Height:=InchesToPoints(0.4), _
                Width:=InchesToPoints(1.35))
            With newBox
                ' Left and Top are required arguments above, but get reset when we change the position so we have to set them again here
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionLeftMarginArea
                .Left = InchesToPoints(0.13)
                .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
                .Top = 0
                .Line.Visible = False
                .LockAnchor = True
                .TextFrame.TextRange.Text = strStyle
                
                With .WrapFormat
                    .Type = wdWrapSquare
                    .Side = wdWrapBoth
                    .DistanceTop = 0
                    .DistanceBottom = 0
                    .DistanceLeft = 0
                    .DistanceRight = 0
                End With
                
            End With
        End If
        
    Next A
    
    strStatus = "* Adding style names to margin..." & vbNewLine & strStatus

    ' Now open the print dialog so user can print the document.
    sglPercentComplete = 0.97
    strStatus = "* Printing document with style names in  margin..." & vbNewLine & strStatus
    
    Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
    DoEvents
    
    Dialogs(wdDialogFilePrint).Show
    
    ' Cleanup
    sglPercentComplete = 1
    strStatus = strStatus & "* Finishing up..." & vbNewLine
    
    Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
    DoEvents
    
Cleanup:
        ' reset Normal style because I'm not sure if it's sticky or not
        With tempDoc.Styles("Normal")
            .Font.Size = currentSize
            .Font.Name = currentName
            .ParagraphFormat.SpaceAfter = currentSpace
        End With
        ' Close temo doc without saving
        tempDoc.Close wdDoNotSaveChanges
    
    ' Return original document to original template
    CurrentDoc.Activate
    Call AttachTemplateMacro.AttachMe(TemplateName:=currentTemplate)
    
    ' Reset settings to original
    With Application
        .DisplayAlerts = lngOpt
        '.Options.PasteFormatBetweenStyledDocuments = lngPasteStyled
        '.Options.PasteFormatBetweenDocuments = lngPasteFormat
        .DisplayStatusBar = currentStatusBar
        .ScreenUpdating = True
        .ScreenRefresh
    End With
    
End Sub

Private Sub PrintStylesPC()

    ' ===== DEPENDENCIES ==========================================================
    ' Requires modules AttachTemplateMacro (and thus also Macmillan style templates), SharedMacros, and ProgressBar
    ' Before you run this, create a text box with the listed settings below, then select the
    ' text box and go to Insert > Text Box > Save Selection to Text Box Gallery (Word 2013). In the
    ' Create New Building Block dialog that opens, name the Building Block "StyleNames1" and
    ' give it the category "Macmillan." Be sure to save it in the template at the path below.
    
    ' TEXT BOX SETTINGS:
    ' Layout > Position > Horizontal: Absolute position 0.13" to the right of Left Margin
    ' Layout > Position > Vertical: Absolute position 0" below paragraph
    ' Layout > Position > Lock Anchor = True
    ' Layout > Position > Allow overlap = true
    ' Layout > Text wrapping > Wrapping style = Square
    ' Layout > Distance from text = 0" (each side)
    ' Layout > Text wrapping > Text wrap = right only
    ' Layout > Size > Height: Absolute 0.4"
    ' Layout > Size > Width: Absolute 1.35"
    ' Format > Shape Styles > Shape Outline: No outline
    
    ' ===== LIMITATIONS ==========================================================
    ' "Normal" can't be in-use as a style in the document
    ' If total margin size (left + right) is < 2 " paragraphs will reflow
    ' Doesn't work for endnotes/footnotes (can't add a drawing object to EN/FNs)
    ' Doesn't work on tables -- breaks the whole macro
    

    ' ======= Run startup checks ========
    ' True means a check failed (e.g., doc protection on)
    If StartupSettings = True Then
        Call Cleanup
        Exit Sub
    End If
    
    
    ' ===== Progress Bar ========================
    Dim sglPercentComplete As Single
    Dim strStatus As String
    Dim strTitle As String
           
    strTitle = "Print Styles in Margin"
    sglPercentComplete = 0.01
    strStatus = "* Getting started..."
    
    Dim objProgressPrint As ProgressBar
    Set objProgressPrint = New ProgressBar  ' Triggers Initialize event, which uses Show methond
    objProgressPrint.Title = strTitle

    Call UpdateBarAndWait(Bar:=objProgressPrint, Status:=strStatus, Percent:=sglPercentComplete)

    
    ' ===== Copy and Paste into a new doc ===================
    ' Paste below throws an alert that too many styles are being copied, so turn off alerts
    ' Also make sure paste settings maintain styles
    Dim lngOpt As Long
    Dim lngPasteStyled As Long
    Dim lngPasteFormat As Long
    
    With Application
    ' record current settings to reset in Cleanup
        lngOpt = .DisplayAlerts
        lngPasteStyled = .Options.PasteFormatBetweenStyledDocuments ' not available on Mac
        lngPasteFormat = .Options.PasteFormatBetweenDocuments  ' not available on Mac
        ' need to display messages because if original template and current template don't match exactly, will
        ' still get the alert that there are too many styles to copy and we need to select the option that
        ' is NOT the default
        .DisplayAlerts = wdAlertsAll
        .Options.PasteFormatBetweenStyledDocuments = wdKeepSourceFormatting
        .Options.PasteFormatBetweenDocuments = wdKeepSourceFormatting
    End With
    
    ' ===== Create new version of this document to manipulate ============
    sglPercentComplete = 0.03
    strStatus = "* Creating dupe document to tag with style names..." & vbNewLine & strStatus
    
    Call UpdateBarAndWait(Bar:=objProgressPrint, Status:=strStatus, Percent:=sglPercentComplete)
    
    ' Copy the text of the document into a new document, so we don't screw up the original
    ' Needs to have the BoundMS template attached before copying so the styles match
    ' the new document later, or it won't copy any styles
    Dim currentTemplate As String
    Dim CurrentDoc As Document
    Set CurrentDoc = ActiveDocument
    ' Record current template
    currentTemplate = CurrentDoc.AttachedTemplate
    
    ' Attach BoundMS template to original doc, then copy contents
    Call AttachTemplateMacro.zz_AttachBoundMSTemplate
    CurrentDoc.StoryRanges(wdMainTextStory).Copy
    
    Dim tempDoc As Document
    ' Create a new document
    Set tempDoc = Documents.Add(Visible:=False)
    ' Add Macmillan styles with no color guides (because if we don't add them,
    ' we get an error that there are too many styles to paste and it just pastes
    ' all with Normal style)
    tempDoc.Activate
    Call AttachTemplateMacro.zz_AttachBoundMSTemplate
    
    'If the template isn't EXACTLY the same (e.g., document was originally styled with an earlier version)
    'you'll still get the error that there are too many styles, so send 'n' to choose "No" from the alert
    ' Not the best solution because if they error isn't thrown it still sends the key and "n" gets typed in doc
    SendKeys "n"
    tempDoc.Content.PasteSpecial DataType:=wdPasteHTML ' maintains styles
        
    ' ===== Set margins =================
    sglPercentComplete = 0.05
    strStatus = "* Adjusting margins to fit style names..." & vbNewLine & strStatus

    Call UpdateBarAndWait(Bar:=objProgressPrint, Status:=strStatus, Percent:=sglPercentComplete)
    
    ' if possible, we want the total margin size to stay the same
    ' so that the paragraphs don't reflow
    ' note margins are in points, 72 pts = 1 inch
    Dim currentLeft As Long
    Dim currentRight As Long
    Dim currentTotal As Long
    
    With tempDoc.PageSetup
        currentLeft = .LeftMargin
        currentRight = .RightMargin
        currentTotal = currentLeft + currentRight
        .LeftMargin = 108   ' 1.5 inches
            If currentTotal >= 144 Then     ' 2 inches
                .RightMargin = currentTotal - 108   ' 1.5 inches
            Else
                .RightMargin = 36   '.5 inches (minimum right margin)
            End If
    End With
    
    ' ==== Change Normal style formatting (since it will define the text boxes) =====
    sglPercentComplete = 0.07
    strStatus = "* Setting format for style names..." & vbNewLine & strStatus
    
    Call UpdateBarAndWait(Bar:=objProgressPrint, Status:=strStatus, Percent:=sglPercentComplete)
        
    ' But save settings first and then change back -- are these settings sticky?
    Dim currentSize As Long
    Dim currentName As String
    Dim currentSpace As Long
    
    With tempDoc.Styles("Normal")
        currentSize = .Font.Size
        currentName = .Font.Name
        currentSpace = .ParagraphFormat.SpaceAfter
        .Font.Size = 7
        .Font.Name = "Calibri"
        .ParagraphFormat.SpaceAfter = 0
    End With
    
    ' ==== Add style names in margin ==================================
    Dim strPath As String
    Dim objTemplate As Template
    Dim objBB As BuildingBlock
    Dim strStyle As String
    Dim lngTextBoxes As Long
    Dim activeParas As Long
    Dim A As Long
    
    ' This is the template where the building block is saved
    strPath = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate\MacmillanGT.dotm"
    If IsItThere(strPath) = True Then
        Set objTemplate = Templates(strPath)
    Else
        MsgBox "I can't find the Macmillan template, sorry."
        GoTo FinishUp
    End If
    
    ' Access the building block through the type and category
    ' NOTE the text box building block has to already be created in the template.
    Set objBB = objTemplate.BuildingBlockTypes(wdTypeTextBox).Categories("Macmillan").BuildingBlocks("StyleNames1")
    
    ' Count the number of current text boxes etc., because the index number of the new ones
    ' will be offset by that amount
    Dim strStatusLoop As String
    Dim lngBoxCount As Long
    lngTextBoxes = tempDoc.Shapes.Count
    activeParas = tempDoc.Paragraphs.Count
    
    For A = 1 To activeParas
        If A Mod 50 = 0 Then
            sglPercentComplete = Round((((A / activeParas) * 0.85) + 0.1), 2)
            strStatusLoop = "* Adding style names to paragraph " & A & " of " & activeParas & "..." & vbNewLine & strStatus

            Call UpdateBarAndWait(Bar:=objProgressPrint, Status:=strStatus, Percent:=sglPercentComplete)
            
            'SecondsElapsed = Round(Timer - StartTime, 2)
            'Debug.Print "Paragraph " & a & " in " & SecondsElapsed & " seconds"
        End If
    
        tempDoc.Paragraphs(A).Range.Select
        strStyle = Selection.Style
        
        ' Don't label Text Std, to save time
        If strStyle <> "Text - Standard (tx)" Then
            lngBoxCount = lngBoxCount + 1
            Selection.Collapse Direction:=wdCollapseStart
            objBB.Insert Where:=Selection.Range        ' works on PC, not on Mac
            tempDoc.Shapes(lngBoxCount).TextFrame.TextRange.Text = strStyle
        End If
        
    Next A
    
    strStatus = "* Adding style names to margin..." & vbNewLine & strStatus

    ' Now open the print dialog so user can print the document.
    sglPercentComplete = 0.97
    strStatus = strStatus & "* Printing document with style names in  margin..." & vbNewLine
    
    Call UpdateBarAndWait(Bar:=objProgressPrint, Status:=strStatus, Percent:=sglPercentComplete)
    
    Dialogs(wdDialogFilePrint).Show
    
    ' Cleanup
    sglPercentComplete = 1
    strStatus = strStatus & "* Finishing up..." & vbNewLine

    Call UpdateBarAndWait(Bar:=objProgressPrint, Status:=strStatus, Percent:=sglPercentComplete)
    
FinishUp:
        ' reset Normal style because I'm not sure if it's sticky or not
        With tempDoc.Styles("Normal")
            .Font.Size = currentSize
            .Font.Name = currentName
            .ParagraphFormat.SpaceAfter = currentSpace
        End With
        ' Close temo doc without saving
        tempDoc.Close wdDoNotSaveChanges
    
    ' Return original document to original template
    CurrentDoc.Activate
    Call AttachTemplateMacro.AttachMe(TemplateName:=currentTemplate)
    
    ' Reset settings to original
    With Application
        .DisplayAlerts = lngOpt
        .Options.PasteFormatBetweenStyledDocuments = lngPasteStyled
        .Options.PasteFormatBetweenDocuments = lngPasteFormat
    End With
    
    Call Cleanup
    Unload objProgressPrint

End Sub

