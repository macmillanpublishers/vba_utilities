VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Document_Open()
' by Erica Warren -- erica.warren@macmillan.com
' removes the Macmillan style templates, macro templates, and assorted files from user's computer
' Note that on Mac you need admin priviledges to delete anything from the Startup directory
    
    ' Alert the user that we are going to uninstall the files
    Dim strMessage As String
    strMessage = "Click OK if you would like to REMOVE the Macmillan Tools tab and styles." & _
                    vbNewLine & vbNewLine & _
                    "Click CANCEL to exit without removing."
                    
    If MsgBox(strMessage, vbOKCancel, "Uninstall Macmillan Tools?") = vbCancel Then
        ActiveDocument.Close (wdDoNotSaveChanges)
        Exit Sub        ' don't really need this since macro will stop when we close the doc
    End If
    
    ' Templates in STARTUP are loaded as global add-ins; unload here before deleting
    AddIns.Unload RemoveFromList:=True
    
    ' Add a step to close all open documents except this one? Because we can't delete style templates
    ' if they are attached to an open document I'm pretty sure
    Call CloseOpenDocs
    
    ' Get paths to the files we need to remove
    Dim strMacUser As String
    Dim strStyleDir As String
    Dim strStartupDir As String
    
    strStartupDir = Application.StartupPath
    
    #If Mac Then
        strMacUser = MacScript("tell application " & Chr(34) & "System Events" & Chr(34) & Chr(13) & _
                "return (name of current user)" & Chr(13) & "end tell")
        strStyleDir = "Macintosh HD/Users/" & strMacUser & "/Documents/MacmillanStyleTemplate"
    #Else
        strStyleDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate"
    #End If
    
    On Error GoTo ErrorHandler
    
    ' Start deletin'
    Dim strLogDir As String
    strLogDir = strStyleDir & Application.PathSeparator & "log"
    
    ' Clear out the log directory
    If IsItThere(Path:=strLogDir) = True Then
        #If Mac Then
            Call DeleteFilesInFolderMac(FolderWithFiles:=strLogDir)
        #Else
            Kill strLogDir & Application.PathSeparator & "*.*"
        #End If
        ' Delete the log folder itself
        RmDir strLogDir
    End If
    
    ' Clear out Style Template directory
    If IsItThere(Path:=strStyleDir) = True Then
        #If Mac Then
            Call DeleteFilesInFolderMac(FolderWithFiles:=strStyleDir)
        #Else
            Kill strStyleDir & Application.PathSeparator & "*.*"
        #End If
        ' Delete the style folder itself
        Debug.Print "strStyleDir = " & strStyleDir
        'Dir ("C:\")
        RmDir strStyleDir
    End If
    
    ' Clear out only our files from Startup
    If IsItThere(strStartupDir) = True Then
        ' specify the files to delete; may have other non-Macmillan files in Startup
        ' also delete both the template and it's version text file
        #If Mac Then
            Call DeleteFilesInFolderMac(FolderWithFiles:=strStartupDir, FileName:="MacmillanGT.*")
            Call DeleteFilesInFolderMac(FolderWithFiles:=strStartupDir, FileName:="GtUpdater.*")
        #Else
            Kill strStartupDir & Application.PathSeparator & "MacmillanGT.*"
            Kill strStartupDir & Application.PathSeparator & "GtUpdater.*"
        #End If
        
    End If
    
    If IsItThere(strStyleDir) = False Then
        MsgBox "The Macmillan Tools template has been removed from your computer."
    End If
    

    On Error GoTo 0
    'ActiveDocument.Close (wdDoNotSaveChanges)
    Exit Sub
    
ErrorHandler:
    If Err.Number = 53 Then     ' "File not found"; we're trying to delete, so just keep going
        Resume Next
    Else
        Debug.Print Err.Number & " " & Err.Description
    End If
    
End Sub


Private Sub DeleteFilesInFolderMac(FolderWithFiles As String, Optional FileName As String = "*")
    ' From http://www.rondebruin.nl/mac/mac012.htm with a few tweaks
    ' Mac 2011 can't use * as wildcard in Kill so we'll do it this way
    ' Note FileName can include * for wildcards; leave out to delete all files in directory

    Dim scriptToRun As String

    If Right(FolderWithFiles, 1) <> ":" Then
        FolderWithFiles = FolderWithFiles & ":"
    End If
    
    scriptToRun = scriptToRun & "tell application " & Chr(34) & _
                  "Finder" & Chr(34) & Chr(13)
    scriptToRun = scriptToRun & _
               "do shell script ""rm "" & quoted form of posix path of " & _
               Chr(34) & FolderWithFiles & """ & " & Chr(34) & FileName & Chr(34) & Chr(13)
    scriptToRun = scriptToRun & "end tell"

    On Error Resume Next
    MacScript (scriptToRun)
    On Error GoTo 0
End Sub



