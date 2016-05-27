Attribute VB_Name = "AttachTemplateMacro"
Option Explicit
'created by Erica Warren - erica.warren@macmillan.com
' ======== PURPOSE =================
' Attaches assorted templates with custom Macmillan styles to the current document

' ======== DEPENDENCIES ============
' 1. Requires SharedMacros module be installed in the same template
' 2. Requires the macmillan style templates be saved in the correct directories
'    that were nstalled on user's computer with Installer file or updated from MacmillanGT.dotm


Sub zz_AttachStyleTemplate()
    Call AttachMe("macmillan.dotm")
End Sub

Sub zz_AttachBoundMSTemplate()
    Call AttachMe("macmillan_NoColor.dotm")
End Sub

Sub zz_AttachCoverTemplate()
    Call AttachMe("macmillan_CoverCopy.dotm")
End Sub

Sub AttachMe(TemplateName As String)
'Attaches a style template from the MacmillanStyleTemplate directory

    Dim currentUser As String
    Dim myFile As String
    Dim strMacDocs As String
    Dim strMacStyleDir As String
        
    'Set template path according to OS
    #If Mac Then
        strMacDocs = MacScript("return (path to documents folder) as string")
        strMacStyleDir = strMacDocs & "MacmillanStyleTemplate"
        myFile = strMacStyleDir & Application.PathSeparator & TemplateName
    #Else
        myFile = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate\" & TemplateName
    #End If
        
    ' Can't attach template to another template, so
    If IsTemplate(ActiveDocument) = False Then
        'Check that file exists
        If IsItThere(myFile) = True Then
        
            'Apply template with Styles
            With ActiveDocument
                .UpdateStylesOnOpen = True
                .AttachedTemplate = myFile
            End With
        Else
            MsgBox "That style template doesn't seem to exist." & vbNewLine & vbNewLine & _
                    "Install the Macmillan Style Template and try again, or contact workflows@macmillan.com for assistance.", _
                    vbCritical, "Oh no!"
        End If
    End If
    
End Sub

Private Function IsTemplate(ByVal objDoc As Document) As Boolean
  Select Case objDoc.SaveFormat
    Case wdFormatTemplate, wdFormatDocument97, _
         wdFormatXMLTemplate, wdFormatXMLTemplateMacroEnabled, _
         wdFormatFlatXMLTemplate, wdFormatFlatXMLTemplateMacroEnabled
      IsTemplate = True
    Case Else
      IsTemplate = False
  End Select
End Function

