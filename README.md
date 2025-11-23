
# Dynamic Sampling VBA Macro

## Overview
This VBA macro dynamically creates a sample from a large dataset in Excel while ensuring:
- The sample size does not exceed **999 rows**.
- Sampling is spread across the entire dataset (top to bottom).
- Handles edge cases where unique categories are very large.
- Preserves original logic by adjusting step size dynamically.

## Features
- **Dynamic Step Size Adjustment**: Automatically increases step size until sample size ≤ 999.
- **Category Awareness**: Builds a dictionary of unique categories for potential proportional sampling.
- **No Manual Reruns**: The macro self-adjusts for any dataset size.

## How It Works
1. User specifies the column containing categories.
2. The macro calculates the total rows and unique categories.
3. It iteratively adjusts the step size until the sample size is within the limit.
4. Colors sampled rows and deletes non-sampled rows.

## Usage
1. Open Excel and press `Alt + F11` to open the VBA editor.
2. Insert a new module and paste the code from `RandomizeDynamicSample.bas`.
3. Run the macro `RandomizeDynamicSample`.
4. Enter the column letter when prompted.

## Example
- Total rows: 13,429
- Unique categories: 3,022
- Max sample: 999 rows
- Final step size: dynamically calculated (e.g., 14)

## License
Free to use and modify.

## 📄 Code Highlights

```vba

Sub RandomizeDynamicSample()
    Dim LastRow As Long, Rng As Range, List As Object
    Dim lTotalRPGs As Long
    Dim sUserCol As String
    Dim shName As String
    Dim StepSize As Long, SampleCount As Long
    Dim MaxSample As Long
    Dim i As Long

    ' User input for column
    sUserCol = Application.InputBox("Enter the RPG column (letter)", "Select column", "A", , , , , 2)
    shName = ActiveSheet.Name

    ' Find last row
    LastRow = Cells(Rows.Count, sUserCol).End(xlUp).Row

    ' Build dictionary of unique categories
    Set List = CreateObject("Scripting.Dictionary")
    For Each Rng In Range(sUserCol & "2:" & sUserCol & LastRow)
        If Not List.Exists(Rng.Value) Then List.Add Rng.Value, Nothing
    Next

    lTotalRPGs = List.Count
    MaxSample = 999 ' Maximum allowed sample size

    ' Start with initial step size
    StepSize = 1

    ' Dynamically adjust step size until sample size <= MaxSample
    Do
        SampleCount = 0
        For i = 2 To LastRow Step StepSize
            SampleCount = SampleCount + 1
        Next

        If SampleCount > MaxSample Then
            StepSize = StepSize + 1
        End If
    Loop While SampleCount > MaxSample

    ' Clear previous colors
    Worksheets(shName).Cells.Interior.ColorIndex = xlNone

    ' Apply sampling with final StepSize
    For i = 2 To LastRow Step StepSize
        Worksheets(shName).Rows(i).Interior.ColorIndex = 4
    Next

    ' Delete non-sampled rows
    For i = LastRow To 2 Step -1
        If Worksheets(shName).Rows(i).Interior.ColorIndex <> 4 Then
            Worksheets(shName).Rows(i).Delete
        End If
    Next

    Worksheets(shName).Cells(1, 1).Select

    MsgBox "Sampling complete. Total sampled rows: " & SampleCount & vbCrLf & _
           "Step size used: " & StepSize, vbInformation
End Sub

