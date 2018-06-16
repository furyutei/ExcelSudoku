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
    Dim IsProtected As Boolean
    
    Set ObjectSudoku = New ClassSudoku
    
    IsProtected = ActiveSheet.ProtectContents
    If IsProtected Then ActiveSheet.Unprotect
    
    With ObjectSudoku
        .ScreenUpdate = ScreenUpdate
        .Logging = Logging
        
        Debug.Print "[ClassSudoku Version " & .Version & "]"
        
        Call ClearSudokuResult
        
        ' ���Ɩ�菉�������Ó����`�F�b�N
        SudokuRange.Font.Color = vbBlack ' �F(vbBlue)�̃Z���̓N���A����Ă��܂��̂ŁA�ύX���Ă���
        If Not .ResetSudokuRange(SudokuRange, TrialCellColor:=vbBlue) Then
            MsgBox "�s���Ȗ��"
            GoTo ExitSub
        End If
        
        ' ���Ɖ�Ǐ���
        Call .TrySudoku(SudokuRange)
        
        ' ���ʎ擾���\��
        Set Result = .LastResult
        TryCounter = Result.Item("TryCounter")
        ElapsedTimeString = Result.Item("ElapsedTimeString")
        
        Debug.Print "����: " & TryCounter & "�񎎍s�E" & ElapsedTimeString & "�b�o��"
    
        If Logging Then
            StageLogLength = Result.Item("StageLogLength")
            If 0 < StageLogLength Then
                ' [���O���e] �X�e�[�W�ԍ�(�����܂����}�X�̐�), �s�ԍ�, ��ԍ�, �ݒ�l(����)
                Range(Cells(1, "U"), Cells(StageLogLength, "X")).Value = Result.Item("StageLog")
            End If
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
    
ExitSub:
    If IsProtected Then ActiveSheet.Protect
End Sub

' ���Ɖ𓚃N���A
Sub ResetSudoku()
    Dim ObjectSudoku As ClassSudoku
    Dim IsProtected As Boolean
    
    Set ObjectSudoku = New ClassSudoku
    
    IsProtected = ActiveSheet.ProtectContents
    If IsProtected Then ActiveSheet.Unprotect
    
    Call ClearSudokuResult
    Call ObjectSudoku.ResetSudokuRange(SudokuRange, TrialCellColor:=vbBlue)
    
    If IsProtected Then ActiveSheet.Protect
End Sub

Private Sub ClearSudokuResult()
    Range("G10").Value = ""
    Range("C10").Value = ""
    
    With Columns("U:X")
        .ClearContents
        If Logging Then
            .Hidden = False
        Else
            .Hidden = True
        End If
    End With
End Sub
