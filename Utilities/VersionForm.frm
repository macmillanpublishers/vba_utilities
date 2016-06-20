VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VersionForm 
   Caption         =   "All the Versions!!"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2655
   OleObjectBlob   =   "VersionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VersionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private c_blnCancel As Boolean


Public Property Let CancelMe(blnCancel As Boolean)
    c_blnCancel = blnCancel
End Property


Public Property Get CancelMe() As Boolean
    CancelMe = c_blnCancel
End Property


' ###################################
' #         EVENTS                  #
' ###################################

Private Sub cmdChange_Click()
    Me.CancelMe = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.CancelMe = True
    Me.Hide
End Sub

Private Sub VersionForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If closed by any means other than Unload code, set Cancel to True, calling procedures should test and unload
    If CloseMove <> 1 Then
        Me.CancelMe = True
        Me.Hide
    End If
End Sub

' ###########################################
' ##        METHODS                         #
' ###########################################

Public Sub PopulateFormData(IndexNumber As Long, NameOfTemplate As String, OldNumber As String)
    ' Add correct template name to VersionForm before displaying to user
    ' NOTE: Index starts at 1, there are 6 total
    ' FUTURE: make number of frames dynamic
    
    Dim OneFrame As MSForms.Control
    Dim OneLabel As MSForms.Control
    
    For Each OneFrame In Me.Controls
        ' Note, each frame name includes sequential numbers
        If TypeOf OneFrame Is MSForms.Frame And InStr(OneFrame.Name, IndexNumber) Then
            ' Add name of template to frame Caption
            ' Also to Tag property so we can access it later
            OneFrame.Caption = NameOfTemplate
            OneFrame.Tag = NameOfTemplate
            
                For Each OneLabel In OneFrame.Controls
                    If OneLabel.Tag = "Current" Then
                        ' Add current version to label text
                        OneLabel.Caption = OldNumber
                        Exit For
                    End If
                Next OneLabel
            
            Exit For
        End If
    Next OneFrame
    
End Sub


Public Function NewVersion(FrameName As String) As String
    ' Pull new version number text from form based on Tag property of frame
    ' will return vbNullString if nothing was entered
    Dim OneFrame As Control
    Dim OneTextbox As Control
    
    For Each OneFrame In Me.Controls
        ' Tag was added in PopulateFormData method
        If TypeOf OneFrame Is MSForms.Frame And OneFrame.Tag = FrameName Then
            For Each OneTextbox In OneFrame.Controls
                If OneTextbox.Tag = "New" Then
                    NewVersion = OneTextbox.Value
                    Exit For
                End If
            Next OneTextbox
            Exit For
        End If
    Next OneFrame
End Function
