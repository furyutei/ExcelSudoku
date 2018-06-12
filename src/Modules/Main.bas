Attribute VB_Name = "Main"
Option Explicit

Private Const ScreenUpdate As Boolean = False ' True: ��ʍX�V

Private Property Get SudokuRange() As Range
    Set SudokuRange = Range("A1:I9") ' �Ώې��ƍs��(9�~9�Œ�)
End Property

Sub TrySudoku()
    Dim ObjectSudoku As ClassSudoku
    Dim Result As Collection
    Dim TryCounter As Long
    Dim ElapsedTimeString As String
    
    Set ObjectSudoku = New ClassSudoku
    
    ObjectSudoku.ScreenUpdate = ScreenUpdate
    
    ' ���Ɩ�菉�������Ó����`�F�b�N
    If Not ObjectSudoku.ResetSudokuRange(SudokuRange) Then
        MsgBox "�s���Ȗ��"
        Exit Sub
    End If
    
    Range("G10").Value = ""
    Range("C10").Value = ""
    
    ' ���Ɖ�Ǐ���
    Call ObjectSudoku.TrySudoku(SudokuRange)
    
    ' ���ʎ擾���\��
    Set Result = ObjectSudoku.LastResult
    TryCounter = Result.Item("TryCounter")
    ElapsedTimeString = Result.Item("ElapsedTimeString")
    
    Debug.Print "����: " & TryCounter & "�񎎍s�E" & ElapsedTimeString & "�b�o��"

    Range("C10").Value = TryCounter
    Range("G10").Value = ElapsedTimeString
    SudokuRange.Cells(1, 1).Select

    ' ���Ɖ񓚃`�F�b�N
    If ObjectSudoku.CheckSudokuRange(SudokuRange) = 0 Then
        MsgBox "��ǐ���"
    Else
        MsgBox "�����c�H"
    End If
End Sub

Sub ResetSudoku()
    Dim ObjectSudoku As ClassSudoku
    
    Set ObjectSudoku = New ClassSudoku

    Range("G10").Value = ""
    Range("C10").Value = ""
    
    Call ObjectSudoku.ResetSudokuRange(SudokuRange)
End Sub
