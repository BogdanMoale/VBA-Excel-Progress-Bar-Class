# VBA Excel Progress Bar Class

This VBA Excel class module provides functionality to create and manage a customizable progress bar within Excel. The progress bar can be displayed as a standalone form or integrated with Excel's status bar to provide real-time feedback on the progress of a task.

## Features

- **Flexible Configuration**: Customize the progress bar's appearance, such as title, colors, and message display.
- **Integration with Excel**: Optionally show progress updates in Excel's status bar for seamless integration with other Excel operations.
- **Dynamic Progress Updates**: Easily update the progress bar's status and message as tasks are completed.
- **User-Friendly Interface**: Simple API makes it easy to integrate the progress bar into existing VBA projects.

## Usage

1. Import the `ProgressBar` class module into your VBA project.
2. Create an instance of the `ProgressBar` class.
3. Configure the progress bar properties as needed, such as title, colors, and total actions.
4. Use the provided methods to update the progress bar as tasks are completed.
5. Optionally, call the `Complete` method to indicate when all tasks are finished.

## Example

```vba
' Create an instance of the ProgressBar class
Dim progressBar As New ProgressBar

' Set properties
progressBar.Title = "Task Progress"
progressBar.TotalActions = 100

' Show the progress bar
progressBar.ShowBar

' Perform tasks and update progress
For i = 1 To 100
    ' Perform task
    ' Update progress
    progressBar.ActionNumber = i
    progressBar.StatusMessage = "Processing task " & i
    ' Optionally, update Excel's status bar
    progressBar.NextAction
Next i

' Complete the progress bar
progressBar.Complete
