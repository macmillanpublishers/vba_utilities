VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Macro Progress"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9345
   OleObjectBlob   =   "ProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






' By Erica Warren - erica.warren@macmillan.com


' ==== USE ==========================
' Displays a userform progress bar to use while another macro is running
' Is modelss on PC so the calling macro will continue to run while progress bar is displayed

' ##### IMPORTANT NOTES!!!  #####
' Word 2011 for Mac can't display userforms modeless so this won't work, however
' all is not lost; the Increment method below will instead update the status bar on Man,
' but be sure not to use ProgressBar.Show method in calling sub
' (Use Load; the Initialize event will show the userform on PC)

' ALSO! Because this is modeless the calling sub will continue to run; this can cause
' problems if the Increment method hasn't finished yet. So there is a Done property; include a line to set it
' as False before using the Increment method, and then use a loop after the Increment method to test for
' ProgressBar.Done = True before continuing.

' Also also, requires SharedMacros modules kinda, because it includes a very helpful
' procedure called UpdateBarAndWait, which, as it says, updates the progress bar and waits
' for ProgressBar.Done = True before continuing

Option Explicit

'========================Progress bar================================.

'Private variables for storing the Main Properties of the Object

'Assigned to Done Property
Private cDone As Boolean

'Assigned to Caption Property of the Form (ProgressBar).
Private cFormTitle As String

'Assigned to Caption Property of LabelCaption control; displays message to user.
Private cStatusMessage As String

'Assigned to Width property of LabelProgress control.
Private cBarWidth As Single

'Assigned to the Caption property of FrameProgress control.
Private cPercentComplete As Single


'##########################################################################
'Class Events
'##########################################################################

'This Procedure is Run when the Form in Initiated
Private Sub UserForm_Initialize()

Me.Done = False

'Set Default Values for all the Variables
' Do I need these here? Won't they be executed when the Property Let procedures for
' the properties below are called?
cFormTitle = "Progress Bar"
cStatusMessage = "Ready"
cPercentComplete = 0
cBarWidth = cPercentComplete * Me.FrameProgress.Width

Me.Title = cFormTitle
Me.LabelCaption.Caption = cStatusMessage
Me.FrameProgress.Caption = cPercentComplete
Me.LabelProgress.Width = cBarWidth

#If Mac Then
    ' Nada! Can't run a useform modeless on Mac 2011, so a progress bar is useless.
    ' But lucky you, the Increment method will update the Mac Status Bar instead
    ' So go right on ahead with this, just DON'T use ProgressBar.Show method in the calling sub
    ' We'll show it below for PCs only
    Me.Hide ' In case someone uses the Show method without the Load method
    Application.DisplayStatusBar = True
#Else
    Me.Show
#End If

Me.Done = True

End Sub


'##########################################################################
'Properties
'##########################################################################
' A Property for if the current method/event has finished
' Be sure to test this in a loop following each use in calling sub
' Since the useform is modeless, execution will continue and sometimes get funky if the
' userform is not finished updating each time
Public Property Let Done(value As Boolean)
    cDone = value
End Property

Public Property Get Done() As Boolean
    Done = cDone
End Property

'This Procedure is Executed when the Title Property is Set
Public Property Let Title(value As String)
cFormTitle = value
Me.Caption = cFormTitle
End Property

'This Procedure lets the User try to access the Title Property.
Public Property Get Title() As String
Title = cFormTitle
End Property

'-----------------------------------------------------------------------
'This Procedure is Executed when the StatusMessage Property is Set.
Public Property Let StatusMessage(value As String)

cStatusMessage = value
Me.LabelCaption = cStatusMessage

End Property

'This Procedure lets the User try to access the StatusMessage Property.
Public Property Get StatusMessage() As String
StatusMessage = cStatusMessage
End Property

'-------------------------------------------------------------------------
'This Procedure is Executed when the Percent Property is Set.
Public Property Let Percent(value As Single)
cPercentComplete = Format(value * 100, "0")
Me.FrameProgress.Caption = cPercentComplete & "%"
End Property

'This Procedure lets the User try to access the Percent Property.
Public Property Get Percent() As Single
Percent = cPercentComplete
End Property

'-------------------------------------------------------------------------
'This Procedure is Executed when the BarWidth Property is Set.
Public Property Let BarWidth(value As Single)
cBarWidth = value * (Me.FrameProgress.Width - 21)
Me.LabelProgress.Width = cBarWidth
End Property

'This Procedure lets the User try to access the BarWidth Property.
Public Property Get BarWidth() As Single
BarWidth = cBarWidth
End Property

'##########################################################################
'Public Methods
'##########################################################################
Public Sub Increment(sglPctComplete As Single, strStatus As String)

Me.StatusMessage = strStatus
Me.Percent = sglPctComplete
Me.BarWidth = sglPctComplete

#If Mac Then
    Application.StatusBar = Me.Title & " " & Me.Percent & "% complete | " & Me.StatusMessage
    DoEvents
#Else
    Me.Repaint
#End If

Me.Done = True

End Sub

