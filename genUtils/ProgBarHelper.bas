Attribute VB_Name = "ProgBarHelper"
Option Explicit
' Things for the ProgressBar form/class that can't live in that module

Sub UpdateBarAndWait(Bar As ProgressBar, Status As String, Percent As Single)
    ' Updates a progress bar and waits until it's done updating before running more code
    ' Required because progress bar form is modeless and the rest of the code will continue to run
    ' while it updates
    On Error GoTo ErrHandler
    Bar.Done = False
    Bar.Increment Percent, Status
    Do
        DoEvents  ' Allows other macro execution to continue
    Loop Until Bar.Done = True
    Exit Sub
ErrHandler:
    Debug.Print Err.Number & ": " & Err.Description
    Bar.Done = True
End Sub
