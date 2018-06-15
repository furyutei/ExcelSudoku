Attribute VB_Name = "Main"
Option Explicit

Private Const ScreenUpdate As Boolean = False ' True: ��ʍX�V
Private Const Logging As Boolean = False ' True: ���O�擾

Private Property Get SudokuRange() As Range
    Set SudokuRange = Range("A1:I9") ' �Ώې��ƍs��(9�~9�Œ�)
End Property

' ���Ɖ��
Sub TrySudoku()
    Dim ObjectSudoku As ClassSudoku
    Dim Result As Collection
    Dim TryCounter As Long
    Dim ElapsedTimeString As String
    Dim StageLogLength As Long
    
    Set ObjectSudoku = New ClassSudoku
    
    With ObjectSudoku
        .ScreenUpdate = ScreenUpdate
        .Logging = Logging
        
        Debug.Print "[ClassSudoku Version " & .Version & "]"
        
        ' ���Ɩ�菉�������Ó����`�F�b�N
        SudokuRange.Font.Color = vbBlack
        If Not .ResetSudokuRange(SudokuRange) Then
            MsgBox "�s���Ȗ��"
            Exit Sub
        End If
        
        Range("G10").Value = ""
        Range("C10").Value = ""
        Range("U:X").ClearContents
        
        ' ���Ɖ�Ǐ���
        Call .TrySudoku(SudokuRange)
        
        ' ���ʎ擾���\��
        Set Result = .LastResult
        TryCounter = Result.Item("TryCounter")
        ElapsedTimeString = Result.Item("ElapsedTimeString")
        
        Debug.Print "����: " & TryCounter & "�񎎍s�E" & ElapsedTimeString & "�b�o��"
    
        If Logging Then
            StageLogLength = Result.Item("StageLogLength")
            If 0 < StageLogLength Then Range(Cells(1, "U"), Cells(StageLogLength, "X")).Value = Result.Item("StageLog")
        End If
        
        Range("C10").Value = TryCounter
        Range("G10").Value = ElapsedTimeString
        SudokuRange.Cells(1, 1).Select
    
        ' ���Ɖ񓚃`�F�b�N
        If .CheckSudokuRange(SudokuRange) = 0 Then
            MsgBox "��ǐ���"
        Else
            MsgBox "�����c�H"
        End If
    End With
End Sub

' ���Ɖ𓚃N���A
Sub ResetSudoku()
    Dim ObjectSudoku As ClassSudoku
    
    Set ObjectSudoku = New ClassSudoku

    With ObjectSudoku
        Range("G10").Value = ""
        Range("C10").Value = ""
        Range("U:X").ClearContents
    
        Call .ResetSudokuRange(SudokuRange)
    End With
End Sub
