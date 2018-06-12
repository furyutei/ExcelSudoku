Attribute VB_Name = "Try250"
Option Explicit

Type SudokuInfo
    NumberValues(1 To 9, 1 To 9) As Variant
    ResultNumberValues As Variant
    Result As Boolean
End Type

Private Property Get Try250Sheet() As Worksheet
    Set Try250Sheet = Worksheets("Try250")
End Property
 
Private Property Get HomeCell() As Range
    Set HomeCell = Try250Sheet.Range("L1")
End Property

Private Property Get SourceSudokuRange() As Range
    Set SourceSudokuRange = Try250Sheet.Range("B1:J2250")
End Property

Private Property Get ResultSudokuRange() As Range
    Set ResultSudokuRange = Try250Sheet.Range("L1:T2250")
End Property

Private Property Get ResultMarkSudokuRange() As Range
    Set ResultMarkSudokuRange = Try250Sheet.Range("U1:U2250")
End Property

Private Property Get ElapsedCell() As Range
    Set ElapsedCell = Try250Sheet.Range("V11")
End Property

Private Property Get ErrorCounterCell() As Range
    Set ErrorCounterCell = Try250Sheet.Range("W12")
End Property

Sub SudokuTry250()
    Dim ObjectSudoku As ClassSudoku
    Dim SudokuNumber As Long
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim RowOffset As Long
    Dim AllNumberValues As Variant
    Dim SudokuList(1 To 250) As SudokuInfo
    Dim ResultMark As String
    Dim ErrorCounter As Long
    Dim StartTime As Double
    Dim EndTime As Double
    Dim ElapsedTime As Double
    Dim ElapsedTimeString As String
    
    Application.ScreenUpdating = False

    Set ObjectSudoku = New ClassSudoku
    
    With Try250Sheet
        .Activate
        HomeCell.Select
        
        ElapsedCell.Value = ""
        
        AllNumberValues = SourceSudokuRange.Value
 
        For SudokuNumber = 1 To 250
            RowOffset = (SudokuNumber - 1) * 9
            
            With SudokuList(SudokuNumber)
                For RowNumber = 1 To 9
                    For ColumnNumber = 1 To 9
                        .NumberValues(RowNumber, ColumnNumber) = AllNumberValues(RowOffset + RowNumber, ColumnNumber)
                    Next ColumnNumber
                Next RowNumber
            End With
        Next SudokuNumber
 
        StartTime = Timer
 
        For SudokuNumber = 1 To 250
            With SudokuList(SudokuNumber)
                .Result = ObjectSudoku.TrySudokuValues(.NumberValues, .ResultNumberValues)
            End With
        Next
 
        EndTime = Timer
        If EndTime < StartTime Then EndTime = EndTime + 24 * 60 * 60
        ElapsedTime = EndTime - StartTime
        ElapsedTimeString = Format(ElapsedTime, "0.000000")
        
        ErrorCounter = 0
        
        With ResultMarkSudokuRange
            For SudokuNumber = 1 To 250
                RowOffset = (SudokuNumber - 1) * 9
                
                With SudokuList(SudokuNumber)
                    For RowNumber = 1 To 9
                        For ColumnNumber = 1 To 9
                            AllNumberValues(RowOffset + RowNumber, ColumnNumber) = .ResultNumberValues(RowNumber, ColumnNumber)
                        Next ColumnNumber
                    Next RowNumber
                    
                    ResultMark = IIf(.Result, "›", "~")
                    If .Result = False Then
                        ErrorCounter = ErrorCounter + 1
                    End If
                End With
                
                .Cells(RowOffset + 1, 1).Value = ResultMark
            Next SudokuNumber
        End With
        
        ResultSudokuRange.Value = AllNumberValues
        ElapsedCell.Value = ElapsedTimeString
        ErrorCounterCell.Value = ErrorCounter
        
        Debug.Print "Œ‹‰Ê: " & ElapsedTimeString & "•bŒo‰ß"
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub SudokuReset250()
    With Try250Sheet
        .Activate
        HomeCell.Select
        
        ResultSudokuRange.ClearContents
        ResultMarkSudokuRange.ClearContents
        ElapsedCell.Value = ""
        ErrorCounterCell.Value = ""
    End With
End Sub

