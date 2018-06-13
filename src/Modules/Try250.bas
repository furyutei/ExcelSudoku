Attribute VB_Name = "Try250"
Option Explicit

Type SudokuInfo ' �ʐ��Ɩ���`
    NumberValues(1 To 9, 1 To 9) As Variant ' ���
    ResultNumberValues As Variant ' ��
    Result As Boolean ' ��͌���(True: �����AFalse: ���s�j
End Type

Private Property Get Try250Sheet() As Worksheet
    Set Try250Sheet = Worksheets("Try250") ' �Ώۃ��[�N�V�[�g
End Property
 
Private Property Get HomeCell() As Range
    Set HomeCell = Try250Sheet.Range("L1") ' �z�[���Z��
End Property

Private Property Get SourceSudokuRange() As Range
    Set SourceSudokuRange = Try250Sheet.Range("B1:J2250") ' 250�␔�ƍs��i���j
End Property

Private Property Get ResultSudokuRange() As Range
    Set ResultSudokuRange = Try250Sheet.Range("L1:T2250") ' 250�␔�ƍs��i�𓚁j
End Property

Private Property Get ResultMarkSudokuRange() As Range
    Set ResultMarkSudokuRange = Try250Sheet.Range("U1:U2250") ' ��͌��ʕ\����
End Property

Private Property Get ElapsedCell() As Range
    Set ElapsedCell = Try250Sheet.Range("V8") ' �o�ߎ��ԕ\���Z��
End Property

Private Property Get ErrorCounterCell() As Range
    Set ErrorCounterCell = Try250Sheet.Range("W9") ' �G���[�i��͎��s�j���\���Z��
End Property

' ����250��A�����
Sub TrySudoku250()
    Application.ScreenUpdating = False
    
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
                    
                    ResultMark = IIf(.Result, "��", "�~")
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
        
        Debug.Print "����: " & ElapsedTimeString & "�b�o��"
    End With
    
    Application.ScreenUpdating = True
End Sub

' ��͌��ʃ��Z�b�g�i�N���A�j
Sub ResetSudoku250()
    Application.ScreenUpdating = False

    With Try250Sheet
        .Activate
        HomeCell.Select
        
        ResultSudokuRange.ClearContents
        ResultMarkSudokuRange.ClearContents
        ElapsedCell.Value = ""
        ErrorCounterCell.Value = ""
    End With
    
    Application.ScreenUpdating = True
End Sub

' ���Ɩ��t�@�C��(CSV)�ǂݍ���
Sub ReadCsvSudoku250()
    Dim CurrentFolder As String
    Dim SudokuBook As Workbook
    Dim SudokuSheet As Worksheet
    Dim TargetSudokuRange As Range
    Dim CsvFileName As String
    Dim CsvBook As Workbook
    Dim CsvSudokuRange As Range
    
    Set SudokuBook = ActiveWorkbook
    Set SudokuSheet = Try250Sheet
    Set TargetSudokuRange = SourceSudokuRange
    
    CurrentFolder = CurDir
    ChDir SudokuBook.Path & "\"
    
    CsvFileName = Application.GetOpenFilename(FileFilter:="���Ɩ��t�@�C��,*.csv", Title:="���Ɩ��t�@�C��(CSV)�I��")
    
    If CsvFileName = "False" Then
        GoTo ExitSub
    End If
    
    Application.ScreenUpdating = False
    
    Set CsvBook = Workbooks.Open(CsvFileName)
    Set CsvSudokuRange = CsvBook.Worksheets(1).Range(TargetSudokuRange.Address).Offset(RowOffset:=1)
    
    SudokuSheet.Unprotect
    TargetSudokuRange.Value = CsvSudokuRange.Value
    SudokuSheet.Protect
    
    Call CsvBook.Close(SaveChanges:=False)
    SudokuBook.Activate
    
    Call ResetSudoku250
    
    Application.DisplayAlerts = False
    SudokuBook.Save
    Application.DisplayAlerts = True
ExitSub:
    ChDir CurrentFolder
    Application.ScreenUpdating = True
End Sub

' ������
Sub InitializeSudou250()
    If MsgBox("���������܂����H" & vbCrLf & "�����Ƃ̖�肪���ׂč폜����܂�!!", Buttons:=vbYesNo, Title:="�������m�F") = vbNo Then
        GoTo ExitSub
    End If
    
    Call ResetSudoku250
    Application.ScreenUpdating = False
    
    With Try250Sheet
        .Activate
        HomeCell.Select
        
        .Unprotect
        SourceSudokuRange.ClearContents
        .Protect
    End With

    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
ExitSub:
    Application.ScreenUpdating = True
End Sub

