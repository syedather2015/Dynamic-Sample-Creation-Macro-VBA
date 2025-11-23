
Attribute VB_Name = "RandomizeDynamicSample"
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
