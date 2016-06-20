VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'=======================================
' By Erica Warren - erica.warren@macmillan.com
'
'===== USE ==============================
' Sets variable names (file names and directories) for Installer sub in FileInstaller module
'
'===== DEPENDENCIES =====================
' This is part 1 of 2. It calls the Installer sub, which downloads and installs the Macmillan style templates and macros.
' If this is meant to be an installer, it should reside in the ThisDocument module in a .docm file.
' If this is meant to be an updater, it should reside in whatever module calls the checking procedure.

' Requires FileInstaller and SharedMacros modules be installed in the same template

Option Explicit
Option Base 1

Private Sub Document_Open()
    ' ==================================================
    ' For debugging; comment out of production code
    ' MsgBox "I am Doc_Open for Installer and I am running"
    ' ==================================================

    Dim downloadBranch As GitBranch
    Dim blnInstaller As Boolean
    Dim strTemplateName As String
    Dim strAllTheTemplates() As String
    
'==========================================================================================
    'DEFINE ALL YOUR VARIABLES HERE
    
    'Make sure to switch to MASTER before going live!
    downloadBranch = releases
    
    'True = standalone installer file, False = daily update checker already installed
    blnInstaller = True
    
    'Plain English name of what you're installing/updating, should be singular
    strTemplateName = "Macmillan Tools and Style Template"
    
    ' Get yr template paths
    strAllTheTemplates = GetTemplatesList(TemplatesYouWant:=installTemplates)

    ' ====== PART 1: REMOVE LEGACY TOOLS TEMPLATE =======
    ' The main template, MacmillanGT.dotm, used to be saved in Startup but since non-admin users
    ' don't have write access to that folder on Mac, we've moved it and will load from this sub. But need to
    ' delete the old one first
    
    Dim strOldGT As String
    strOldGT = Application.StartupPath & Application.PathSeparator & "MacmillanGT.dotm"

    On Error GoTo ErrHandler
    If IsItThere(strOldGT) = True Then
        AddIns(strOldGT).Installed = False
        Kill strOldGT
    End If

    ' ====== PART 2: CHECK FOR UPDATES ===============
    ' Unload all add-ins first
    Dim a As Long
    
    For a = LBound(strAllTheTemplates()) To UBound(strAllTheTemplates())
        On Error Resume Next        'Error = add-in not available, don't need to uninstall
            AddIns(strAllTheTemplates(a)).Installed = False

        On Error GoTo 0
    Next a
    
    ' Run Installer macro
    'Installer = True if this is for a standalone installation file; False is for a daily update checker

    Call SharedFileInstaller.Installer(DownloadFrom:=downloadBranch, Installer:=blnInstaller, _
        TemplateName:=strTemplateName, TemplatesToInstall:=strAllTheTemplates())
    
    ' Load GtUpdater add-in (which will then load MacGT add-in)
    
    On Error Resume Next        'Error = add-in not available
        AddIns(strAllTheTemplates(1)).Installed = True
    On Error GoTo 0
    
    ThisDocument.Close
    Exit Sub
    
ErrHandler:
    If Err.Number = 75 Then
        Dim strMessage As String
        ' Can we tell if this is a Macmillan install, and prompt them to use Self Service?"
        ' Could we launch Self Service from AppleScript or something?
        strMessage = "Uh oh, looks like you need to re-install the Macmillan template from Self Service or use an account with administrator priviledges."
        MsgBox strMessage, vbOKOnly, "New template needed"
    Else
        MsgBox Err.Number & ": " & Err.Description
    End If
    
End Sub
