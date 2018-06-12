Attribute VB_Name = "Main"
Option Explicit

Private Const ScreenUpdate As Boolean = False ' True: 画面更新

Private Property Get SudokuRange() As Range
    Set SudokuRange = Range("A1:I9") ' 対象数独行列(9×9固定)
End Property

Sub TrySudoku()
    Dim ObjectSudoku As ClassSudoku
    Dim Result As Collection
    Dim TryCounter As Long
    Dim ElapsedTimeString As String
    
    Set ObjectSudoku = New ClassSudoku
    
    ObjectSudoku.ScreenUpdate = ScreenUpdate
    
    ' 数独問題初期化＆妥当性チェック
    If Not ObjectSudoku.ResetSudokuRange(SudokuRange) Then
        MsgBox "不正な問題"
        Exit Sub
    End If
    
    Range("G10").Value = ""
    Range("C10").Value = ""
    
    ' 数独解読処理
    Call ObjectSudoku.TrySudoku(SudokuRange)
    
    ' 結果取得＆表示
    Set Result = ObjectSudoku.LastResult
    TryCounter = Result.Item("TryCounter")
    ElapsedTimeString = Result.Item("ElapsedTimeString")
    
    Debug.Print "結果: " & TryCounter & "回試行・" & ElapsedTimeString & "秒経過"

    Range("C10").Value = TryCounter
    Range("G10").Value = ElapsedTimeString
    SudokuRange.Cells(1, 1).Select

    ' 数独回答チェック
    If ObjectSudoku.CheckSudokuRange(SudokuRange) = 0 Then
        MsgBox "解読成功"
    Else
        MsgBox "あれれ…？"
    End If
End Sub

Sub ResetSudoku()
    Dim ObjectSudoku As ClassSudoku
    
    Set ObjectSudoku = New ClassSudoku

    Range("G10").Value = ""
    Range("C10").Value = ""
    
    Call ObjectSudoku.ResetSudokuRange(SudokuRange)
End Sub
