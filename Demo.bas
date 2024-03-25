Attribute VB_Name = "Demo"
Option Explicit

Sub TestTheBAR()
Attribute TestTheBAR.VB_ProcData.VB_Invoke_Func = "Q\n14"

'Declare Sub Level Variables and Objects
Dim Counter As Long
Dim TotalCount As Long
Dim SomeRange As Range
Dim EachCell As Range

'Initialize the Variables and Objects
TotalCount = 10
Set SomeRange = Sheet1.UsedRange
SomeRange.Clear

'Declare the ProgressBar Object
Dim MyProgressbar As ProgressBar
'Initialize a New Instance of the Progressbars
Set MyProgressbar = New ProgressBar

'Set all the Properties that need to be set before the
'ProgresBar is Shown
With MyProgressbar
    .Title = "Test The Bar" 'Optional
    .ExcelStatusBar = True 'Optional
    .StartColour = rgbMediumSeaGreen 'Optional
    .EndColour = rgbGreen 'Optional
    .TotalActions = TotalCount 'Required
End With

'Try to Access the Properties that were set
'<<The followinf lines are not essential. This is just
'to let you know that you may access the properties if
'you need it in your code.
Debug.Print MyProgressbar.Title
Debug.Print MyProgressbar.TotalActions
Debug.Print MyProgressbar.ActionNumber
Debug.Print MyProgressbar.StatusMessage

'Show the Bar
MyProgressbar.ShowBar 'Critical Line

For Counter = 1 To TotalCount
    For Each EachCell In SomeRange
        On Error Resume Next 'To Avoid the Too Many Formats Error
        EachCell.Interior.Color = RGB(255 * Rnd(Counter), 255 * Rnd(Counter), 255 * Rnd(Counter))
        On Error GoTo 0
    Next EachCell
    'Update the ProgressBar NextAction Method
    MyProgressbar.NextAction "Struggling To Excel " & Counter, True 'First method for animating the bar
Next Counter


'Check if the Override Properties work
'Also, the following block is not essential. It is here just to
'Illustrate the second method for animating the bar
With MyProgressbar
    .ActionNumber = 5 'Second Method for animating the bar
    .StatusMessage = "Override Test"
    'Wait to show the Change
    Application.Wait (Now() + TimeValue("00:00:02"))
    'Set it back to TotalActions so the Complete Method works
    .ActionNumber = MyProgressbar.TotalActions
End With

MyProgressbar.Complete 3

SomeRange.Clear
End Sub


Sub TestTheSubBAR()

'Declare Sub Level Variables and Objects
Dim Counter As Long
Dim SubCounter As Long
Dim TotalCount As Long
Dim SomeRange As Range
Dim EachCell As Range

'Initialize the Variables and Objects
TotalCount = 10

Set SomeRange = Sheet1.UsedRange
SomeRange.Clear

'Declare the ProgressBar Objects
Dim MainBar As ProgressBar
Dim SubBar As ProgressBar

'Initialize a New Instance of the Progressbars
Set MainBar = New ProgressBar
Set SubBar = New ProgressBar

'Set all the Properties that need to be set before the
'ProgresBar is Shown
With MainBar
    .Title = "Main Bar"
    .ExcelStatusBar = True
    .StartColour = rgbGreen
    .EndColour = rgbRed
    .TotalActions = TotalCount
End With

With SubBar
    .Title = "Sub Bar"
    .ExcelStatusBar = True
    .StartColour = rgbGreen
    .EndColour = rgbRed
End With

'Show the Bar
MainBar.ShowBar

For Counter = 1 To TotalCount
    'Set the Sub Counter = 0
    SubCounter = 0
    'Set the total actions property
    SubBar.TotalActions = SomeRange.Count
    'Show the Sub bar
    SubBar.ShowBar
    'Move the Second bar below the main Bar
    SubBar.Top = MainBar.Top + MainBar.Height + 10
    SubBar.Left = MainBar.Left
    For Each EachCell In SomeRange
        On Error Resume Next 'To Avoid the Too Many Formats Error
        EachCell.Interior.Color = RGB(255 * Rnd(Counter), 255 * Rnd(Counter), 255 * Rnd(Counter))
        On Error GoTo 0
        SubCounter = SubCounter + 1
        SubBar.NextAction "Colouring Cells", True
    Next EachCell
    'Update the ProgressBar NextAction Method
    MainBar.NextAction "Struggling To Excel " & Counter, True
    SubBar.Terminate
Next Counter


MainBar.Complete 3

SomeRange.Clear
End Sub
