Attribute VB_Name = "Demo2"
Option Explicit

'Declare the ProgressBar Objects
Public MainBar As ProgressBar
Public SubBar As ProgressBar

Sub CallTheMainBar()

'Declare Sub Level Variables and Objects
Dim Counter As Long
Dim TotalCount As Long


'Initialize the Variables and Objects
TotalCount = 10


'Initialize a New Instance of the Progressbars
Set MainBar = New ProgressBar


'Set all the Properties that need to be set before the
'ProgresBar is Shown
With MainBar
    .Title = "Main Bar"
    .ExcelStatusBar = True
    .StartColour = rgbGreen
    .EndColour = rgbRed
    .TotalActions = TotalCount
End With

'Show the Bar
MainBar.ShowBar

For Counter = 1 To TotalCount
    'Call other Procedures
    CallTheSubBar
Next Counter

MainBar.Complete 3

End Sub

Sub CallTheSubBar()

Dim Counter As Long
Dim SomeRange As Range
Dim EachCell As Range


Set SubBar = New ProgressBar
Set SomeRange = Sheet1.UsedRange
SomeRange.Clear

With SubBar
    .Title = "Sub Bar"
    .ExcelStatusBar = True
    .StartColour = rgbGreen
    .EndColour = rgbRed
End With

'Set the Sub Counter = 0
Counter = 0
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
    Counter = Counter + 1
    SubBar.NextAction "Colouring Cells", True
Next EachCell
SubBar.Terminate

'Update the ProgressBar NextAction Method - Just to check if This Sub Can access the main bar
MainBar.NextAction "Struggling To Excel " & Counter, True

SomeRange.Clear

End Sub

