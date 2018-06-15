Attribute VB_Name = "Main"
Option Explicit

Private Const ScreenUpdate As Boolean = False ' True: 画面更新
Private Const Logging As Boolean = False ' True: ログ取得

Private Property Get SudokuRange() As Range
    Set SudokuRange = Range("A1:I9") ' 対象数独行列(9×9固定)
End Property

' 数独解析
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
        
        ' 数独問題初期化＆妥当性チェック
        SudokuRange.Font.Color = vbBlack
        If Not .ResetSudokuRange(SudokuRange) Then
            MsgBox "不正な問題"
            Exit Sub
        End If
        
        Range("G10").Value = ""
        Range("C10").Value = ""
        Range("U:X").ClearContents
        
        ' 数独解読処理
        Call .TrySudoku(SudokuRange)
        
        ' 結果取得＆表示
        Set Result = .LastResult
        TryCounter = Result.Item("TryCounter")
        ElapsedTimeString = Result.Item("ElapsedTimeString")
        
        Debug.Print "結果: " & TryCounter & "回試行・" & ElapsedTimeString & "秒経過"
    
        If Logging Then
            StageLogLength = Result.Item("StageLogLength")
            If 0 < StageLogLength Then Range(Cells(1, "U"), Cells(StageLogLength, "X")).Value = Result.Item("StageLog")
        End If
        
        Range("C10").Value = TryCounter
        Range("G10").Value = ElapsedTimeString
        SudokuRange.Cells(1, 1).Select
    
        ' 数独回答チェック
        If .CheckSudokuRange(SudokuRange) = 0 Then
            MsgBox "解読成功"
        Else
            MsgBox "あれれ…？"
        End If
    End With
End Sub

' 数独解答クリア
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
