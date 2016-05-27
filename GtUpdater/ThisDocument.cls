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
'===== USE ===============================
' Sets variable names (file names and directories) for Installer sub in FileInstaller module.
'
'===== DEPENDENCIES ========================
' This is part 1 of 2. It calls the Installer sub, which downloads and installs the Macmillan style templates and macros.
' If this is meant to be an installer, it should be called "Document_Open()" and reside in the ThisDocument module in a .docm file.
' If this is meant to be an updater, it should be called "AutoExec()" and reside in the ThisDocument module in a .dotm file,
' and be saved in the STARTUP directory.

' This requires that the SharedMacros module be saved in the same template

Option Explicit
Option Base 1

Sub AutoExec()

    ' ======= For testing / debugging ==============
    ' MsgBox "I am NEW GtUpdater and I am running!"
    ' ==============================================

    Dim downloadBranch As GitBranch
    Dim blnInstaller As Boolean
    Dim strTemplateName As String
    Dim strTemplateList() As String

'==========================================================================================
    'DEFINE ALL YOUR VARIABLES HERE

    'Make sure to switch to MASTER before going live!
    downloadBranch = releases

    'True = standalone installer file, False = daily update checker already installed
    blnInstaller = False

    'Plain English name of what you're installing/updating, should be singular
    strTemplateName = "Macmillan Tools Template"

    strTemplateList = GetTemplatesList(TemplatesYouWant:=toolsTemplates)


    ' ====== PART 1: REMOVE LEGACY TOOLS TEMPLATE =======
    ' The main template, MacmillanGT.dotm, used to be saved in Startup but since non-admin users
    ' don't have write access to that folder on Mac, we've moved it and will load from this sub. But need to
    ' delete the old one first
    
    ' Um let's figure out a better way to get this later maybe
    Dim strOldGT As String
    strOldGT = Application.StartupPath & Application.PathSeparator & "MacmillanGT.dotm"
    
    On Error GoTo ErrHandler
    If IsItThere(strOldGT) = True Then
        AddIns(strOldGT).Installed = False
        Kill strOldGT
    End If

    ' ====== PART 2: CHECK FOR UPDATES ===============
    'Installer = True if this is for a standalone installation file; False is for a daily update checker
    Call SharedFileInstaller.Installer(DownloadFrom:=downloadBranch, Installer:=blnInstaller, _
        TemplateName:=strTemplateName, TemplatesToInstall:=strTemplateList())

    ' ====== PART 3: LOAD MAIN TEMPLATE ==============
    Dim strNewGT As String
    strNewGT = strTemplateList(1)
    'Debug.Print strNewGT

    If IsItThere(strNewGT) = True Then
        ' Have to add Template before loading if not in Startup
        ' Could do Install:=True with Add method, but it somehow broke everything
        AddIns.Add (strNewGT)
        AddIns(strNewGT).Installed = True
    Else
        Dim strMessage As String
            strMessage = "Wow, this is really embarrassing but something went wrong with your Macmillan Tools template. Please send an email to " & _
            "workflows@macmillan.com and let them know that something is broken!"
        MsgBox strMessage, vbOKOnly, "We're So Sorry"
    End If

    On Error GoTo 0
    Exit Sub

ErrHandler:
    If Err.Number = 75 Then     ' Change to whatever the actual error is when try to kill file
        ' Can we tell if this is a Macmillan install, and prompt them to use Self Service?"
        ' Could we launch Self Service from AppleScript or something?
        strMessage = "Uh oh, looks like you need to re-install the Macmillan template from Self Service or use an account with administrator priviledges."
        MsgBox strMessage, vbOKOnly, "New template needed"
    Else
        MsgBox Err.Number & ": " & Err.Description
    End If


End Sub
