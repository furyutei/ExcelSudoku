VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassSudoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'■ 共通変数 {
Private Const CurrentVersion As String = "0.0.1.6"

' オプション用
Private CurrentScreenUpdate As Boolean ' True: 画面更新
Private StageLogIsEnabled As Boolean ' True: ログ取得
Private CheckInDescendingOrder As Boolean ' True: マスに入る候補が複数ある場合、降順にチェック

' 行列←→グループ座標変換用
Private Type PostionToGroupItem
    GroupNumber As Long
    GroupItemNumber As Long
End Type

Private Type GroupToPositionItem
    RowNumber As Long
    ColumnNumber As Long
End Type

Private GroupMapInitialized As Boolean
Private PositionToGroupMap(1 To 9, 1 To 9) As PostionToGroupItem
Private GroupToPositionMap(1 To 9, 1 To 9) As GroupToPositionItem

' ビット演算用
Private Type NumberValueList
    Length As Long
    NumberValues(1 To 9) As Long
End Type

Private BitOperationInitialized As Boolean
Private Bit() As Long
Private BitAll As Long
Private BitmapToNumberValues() As NumberValueList

' 試行履歴用
Private Type NumberMaps
    RemainNumberMap As Long
End Type

Private Type MapInfo
    Matrix(1 To 9, 1 To 9) As NumberMaps
End Type

Private Type StageInfo
    RowNumber As Long
    ColumnNumber As Long
    NumberValue As Long
    MapInfo As MapInfo
End Type

Private StageHistoryInitialized As Boolean
Private StageHistoryLength As Long
Private StageHistory(0 To 81) As StageInfo

Private StageLogLength As Long
Private StageLog() As String

' 数独結果格納用
Private SudokuRange As Range ' 数独用セル範囲
Private NumberMatrix() As Long ' 数独行列格納用
Private TryResult As Boolean ' 試行結果
Private TryCounter As Long ' 試行回数
Private StartTime As Double ' 試行開始タイマ値
Private EndTime As Double ' 試行終了タイマ値
Private ElapsedTime As Double ' 経過時間
Private ElapsedTimeString As String ' 経過時間(Formatted)
Private LastErrorInfos() As Variant ' エラー情報
'}

'■プロパティ(Public) {
' バージョン取得用
Public Property Get Version() As String
    Version = CurrentVersion
End Property

' 画面更新有効／無効切替用
Public Property Let ScreenUpdate(ByVal Flag As Boolean)
    CurrentScreenUpdate = Flag
End Property

Public Property Get ScreenUpdate() As Boolean
    ScreenUpdate = CurrentScreenUpdate
End Property

' ログ取得切替用
Public Property Let Logging(ByVal Flag As Boolean)
    StageLogIsEnabled = Flag
End Property

Public Property Get Logging() As Boolean
    Logging = StageLogIsEnabled
End Property

' 数値チェック方法（マスに入る候補のチェック順）切替用
Public Property Let CheckOrderDesc(ByVal Flag As Boolean)
    CheckInDescendingOrder = Flag
End Property

Public Property Get CheckOrderDesc() As Boolean
    CheckOrderDesc = CheckInDescendingOrder
End Property

' 最新解析結果取得用
Public Property Get LastResult() As Collection
    Set LastResult = New Collection
    
    Call LastResult.Add(SudokuRange, "SudokuRange")
    Call LastResult.Add(NumberMatrix, "NumberMatrix")
    Call LastResult.Add(TryResult, "TryResult")
    Call LastResult.Add(TryCounter, "TryCounter")
    Call LastResult.Add(ElapsedTime, "ElapsedTime")
    Call LastResult.Add(StartTime, "StartTime")
    Call LastResult.Add(EndTime, "EndTime")
    Call LastResult.Add(ElapsedTimeString, "ElapsedTimeString")
    Call LastResult.Add(StageLogLength, "StageLogLength")
    Call LastResult.Add(WorksheetFunction.Transpose(StageLog), "StageLog")
    Call LastResult.Add(LastErrorInfos, "ErrorInfos")
End Property
'}

'■ プロパティ(Private) {
Private Property Get CurrentStageNumber() As Long
    CurrentStageNumber = StageHistoryLength
End Property

Private Property Get CurrentStage() As StageInfo
    CurrentStage = StageHistory(StageHistoryLength)
End Property

Private Property Get IsLastStage() As Boolean
    If 81 <= StageHistoryLength Then
        IsLastStage = True
    Else
        IsLastStage = False
    End If
End Property
'}

'■ コンストラクタ／デコンストラクタ {
Private Sub Class_Initialize()
    CurrentScreenUpdate = False
    StageLogIsEnabled = False
    CheckInDescendingOrder = False
    
    Call InitializeBitOperation
    Call InitializeGroupMap
    Call InitializeStageHistory
End Sub

Private Sub Class_Terminate()
End Sub
'}

'■ メソッド(Public) {
' 数独解析（Range指定版）
Public Function TrySudoku(ByRef TargetSudokuRange As Range, Optional ByRef Reset As Boolean = False, Optional ByRef Update As Boolean = True) As Boolean
    Dim RemainCounter As Long
    Dim ErrorInfos() As Variant
    
    If Reset Then Call ResetSudokuRange(TargetSudokuRange)
    
    Set SudokuRange = TargetSudokuRange
    
    StartTime = Timer
    
    TryCounter = 0
    
    RemainCounter = InitializeSudokuMatrix(SudokuRange, ErrorInfos)
    If RemainCounter < 0 Then
        LastErrorInfos = ErrorInfos
        TrySudoku = False
        GoTo ExitFunction
    End If
    
    TryResult = TrySudokuMatrix()
    TrySudoku = TryResult
    
ExitFunction:
    EndTime = Timer
    If EndTime < StartTime Then EndTime = EndTime + 24 * 60 * 60
    
    ElapsedTime = EndTime - StartTime
    ElapsedTimeString = Format(ElapsedTime, "0.0000")
    
    If Update Then SudokuRange.Value = NumberMatrix
End Function

' 数独解析（配列指定版）
Public Function TrySudokuValues(ByRef TargetNumberValues As Variant, ByRef ResultNumberValues As Variant) As Boolean
    Dim RemainCounter As Long
    Dim ErrorInfos() As Variant
    
    Set SudokuRange = Nothing
    
    StartTime = Timer
    
    TryCounter = 0
    
    RemainCounter = InitializeSudokuMatrix(TargetNumberValues, ErrorInfos)
    If RemainCounter < 0 Then
        LastErrorInfos = ErrorInfos
        TrySudokuValues = False
        GoTo ExitFunction
    End If
    
    TryResult = TrySudokuMatrix()
    TrySudokuValues = TryResult
    
ExitFunction:
    EndTime = Timer
    If EndTime < StartTime Then EndTime = EndTime + 24 * 60 * 60
    
    ElapsedTime = EndTime - StartTime
    ElapsedTimeString = Format(ElapsedTime, "0.0000")
    
    ResultNumberValues = NumberMatrix
End Function

' 数独問題(Range)リセット（回答クリア）
' ※試行セル（文字色が青(vbBlue)のセル）はクリアされることに注意
Public Function ResetSudokuRange(ByRef TargetSudokuRange As Range) As Boolean
    Application.ScreenUpdating = False
    Application.CutCopyMode = False

    Dim SudokuRangeValues As Variant
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim NumberMatrix() As Long
    Dim ErrorInfos() As Variant
    Dim ErrorIndex As Long
    
    ReDim NumberMatrix(1 To 9, 1 To 9)
    
    With TargetSudokuRange
        .Cells(1, 1).Select
        
        SudokuRangeValues = .Value
    
        For RowNumber = 1 To 9
            For ColumnNumber = 1 To 9
                If SudokuRangeValues(RowNumber, ColumnNumber) = 0 Or .Cells(RowNumber, ColumnNumber).Font.Color = vbBlue Then
                    NumberMatrix(RowNumber, ColumnNumber) = 0
                    .Cells(RowNumber, ColumnNumber).Value = ""
                    .Cells(RowNumber, ColumnNumber).Font.Color = vbBlue
                Else
                    NumberMatrix(RowNumber, ColumnNumber) = .Cells(RowNumber, ColumnNumber).Value
                    .Cells(RowNumber, ColumnNumber).Font.Color = vbBlack
                End If
                .Cells(RowNumber, ColumnNumber).Interior.ColorIndex = xlNone
            Next ColumnNumber
        Next RowNumber
        
        If 0 <= InitializeSudokuMatrix(NumberMatrix, ErrorInfos, CheckOnly:=True) Then
            ResetSudokuRange = True
        Else
            For ErrorIndex = 1 To UBound(ErrorInfos)
                .Cells(ErrorInfos(ErrorIndex)(0), ErrorInfos(ErrorIndex)(1)).Interior.Color = vbRed
            Next ErrorIndex
            ResetSudokuRange = False
        End If
    End With
    
    Application.ScreenUpdating = True
End Function

' 数独問題(Range)の妥当性チェック
' ※残りの空欄数を返す
' ※エラー（重複）セルがあれば、背景色を赤(vbRed)にする
Public Function CheckSudokuRange(ByRef TargetSudokuRange As Range) As Long
    Application.ScreenUpdating = False
    Application.CutCopyMode = False

    Dim SudokuRangeValues As Variant
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim NumberMatrix() As Long
    Dim ErrorInfos() As Variant
    Dim ErrorIndex As Long
    Dim RemainCounter As Long
    
    ReDim NumberMatrix(1 To 9, 1 To 9)
    
    With TargetSudokuRange
        .Cells(1, 1).Select
        
        SudokuRangeValues = .Value
        
        For RowNumber = 1 To 9
            For ColumnNumber = 1 To 9
                NumberMatrix(RowNumber, ColumnNumber) = SudokuRangeValues(RowNumber, ColumnNumber)
            Next ColumnNumber
        Next RowNumber
        
        RemainCounter = InitializeSudokuMatrix(NumberMatrix, ErrorInfos, CheckOnly:=True)
        
        If RemainCounter < 0 Then
            For ErrorIndex = 1 To UBound(ErrorInfos)
                .Cells(ErrorInfos(ErrorIndex)(0), ErrorInfos(ErrorIndex)(1)).Interior.Color = vbRed
            Next ErrorIndex
        End If
    End With
        
    CheckSudokuRange = RemainCounter
    
    Application.ScreenUpdating = True
End Function
'}

'■ メソッド(Private) {
Private Function InitializeSudokuMatrix(ByRef SudokuRangeValues As Variant, ByRef ErrorInfos() As Variant, Optional CheckOnly As Boolean = False) As Long
    Dim RemainCounter As Long
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim GroupNumber As Long
    Dim GroupItemNumber As Long
    Dim CheckRowNumber As Long
    Dim CheckColumnNumber As Long
    Dim CheckGroupItemNumber As Long
    Dim NumberValue As Long
    Dim CheckNumberValue As Long
    Dim MapInfo As MapInfo
    Dim Index As Long
    Dim RowRemainNumberMap(1 To 9) As Long
    Dim ColumnRemainNumberMap(1 To 9) As Long
    Dim GroupRemainNumberMap(1 To 9) As Long
    Dim RowNumberValueCounters(1 To 9, 1 To 9) As Long
    Dim ColumnNumberValueCounters(1 To 9, 1 To 9) As Long
    Dim GroupNumberValueCounters(1 To 9, 1 To 9) As Long
    
    RemainCounter = 81
    
    ReDim NumberMatrix(1 To 9, 1 To 9)
    ReDim ErrorInfos(0)
    
    For Index = 1 To 9
        RowRemainNumberMap(Index) = BitAll
        ColumnRemainNumberMap(Index) = BitAll
        GroupRemainNumberMap(Index) = BitAll
    Next
    
    Call ResetStageHistory
    MapInfo = StageHistory(0).MapInfo

    For RowNumber = 1 To 9
        For ColumnNumber = 1 To 9
            NumberValue = SudokuRangeValues(RowNumber, ColumnNumber)
            If Not CheckOnly Then NumberMatrix(RowNumber, ColumnNumber) = NumberValue
            
            If NumberValue <> 0 Then
                Call PositionToGroup(RowNumber, ColumnNumber, GroupNumber, GroupItemNumber)
                
                With MapInfo.Matrix(RowNumber, ColumnNumber)
                    .RemainNumberMap = Bit(NumberValue)
                End With
                
                Call BitOff(RowRemainNumberMap(RowNumber), NumberValue)
                Call BitOff(ColumnRemainNumberMap(ColumnNumber), NumberValue)
                Call BitOff(GroupRemainNumberMap(GroupNumber), NumberValue)
                
                RowNumberValueCounters(RowNumber, NumberValue) = RowNumberValueCounters(RowNumber, NumberValue) + 1
                ColumnNumberValueCounters(ColumnNumber, NumberValue) = ColumnNumberValueCounters(ColumnNumber, NumberValue) + 1
                GroupNumberValueCounters(GroupNumber, NumberValue) = GroupNumberValueCounters(GroupNumber, NumberValue) + 1
            End If
        Next ColumnNumber
    Next RowNumber
                
    For RowNumber = 1 To 9
        For ColumnNumber = 1 To 9
            NumberValue = SudokuRangeValues(RowNumber, ColumnNumber)
            
            If NumberValue = 0 Then
                Call PositionToGroup(RowNumber, ColumnNumber, GroupNumber, GroupItemNumber)
                
                With MapInfo.Matrix(RowNumber, ColumnNumber)
                    .RemainNumberMap = RowRemainNumberMap(RowNumber) And ColumnRemainNumberMap(ColumnNumber) And GroupRemainNumberMap(GroupNumber)
                End With
            Else
                If _
                    1 < RowNumberValueCounters(RowNumber, NumberValue) Or _
                    1 < ColumnNumberValueCounters(ColumnNumber, NumberValue) Or _
                    1 < GroupNumberValueCounters(GroupNumber, NumberValue) _
                Then
                    ReDim Preserve ErrorInfos(UBound(ErrorInfos) + 1)
                    ErrorInfos(UBound(ErrorInfos)) = Array(RowNumber, ColumnNumber)
                Else
                    RemainCounter = RemainCounter - 1
                End If
            End If
        Next ColumnNumber
    Next RowNumber
    
    If 0 < UBound(ErrorInfos) Then
        InitializeSudokuMatrix = -1
        Exit Function
    End If
    
    If Not CheckOnly Then
        StageHistoryLength = 81 - RemainCounter
        StageHistory(StageHistoryLength).MapInfo = MapInfo
    End If
    
    InitializeSudokuMatrix = RemainCounter
End Function

Private Function TrySudokuMatrix() As Boolean
    Dim SavedStageNumber As Long
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim GroupNumber As Long
    Dim GroupItemNumber As Long
    Dim LastSetRowNumber As Long
    Dim LastSetColumnNumber As Long
    Dim NumberList As NumberValueList
    Dim MinPoint As Long
    Dim CurrentPoint As Long
    Dim CheckRowNumber As Long
    Dim CheckColumnNumber As Long
    Dim CheckGroupItemNumber As Long
    Dim TryRowNumber As Long
    Dim TryColumnNumber As Long
    Dim TryNumberList As NumberValueList
    Dim Index As Long
    Dim StartIndex As Long
    Dim EndIndex As Long
    Dim StepIndex As Long
    
    If ScreenUpdate Then Range("A1:I9").Value = NumberMatrix
    
    GoSub CheckLastStage
    
    SavedStageNumber = CurrentStageNumber
    
    MinPoint = 9999
    LastSetRowNumber = 10
    LastSetColumnNumber = 10
    
    Do
        For RowNumber = 1 To 9
            For ColumnNumber = 1 To 9
                If RowNumber = LastSetRowNumber And ColumnNumber = LastSetColumnNumber Then GoTo BreakLoop
            
                If NumberMatrix(RowNumber, ColumnNumber) = 0 Then
                    NumberList = GetRemainNumberList(RowNumber, ColumnNumber)
                    
                    With NumberList
                        If .Length < 1 Then
                            ' 候補が無ければ試行は失敗
                            GoTo Failure
                        ElseIf .Length = 1 Then
                            ' 候補がひとつだけのマスを埋める
                            Call PushStage(RowNumber, ColumnNumber, .NumberValues(1))
                            
                            GoSub CheckLastStage
                            
                            MinPoint = 9999
                            LastSetRowNumber = RowNumber
                            LastSetColumnNumber = ColumnNumber
                        Else
                            ' 候補がより少ないマスを記憶
                            CurrentPoint = .Length * 1000
                            
                            If CurrentPoint < MinPoint Then
                                TryRowNumber = RowNumber
                                TryColumnNumber = ColumnNumber
                                MinPoint = CurrentPoint
                            End If
                            
                        End If
                    End With
                End If
            Next ColumnNumber
        Next RowNumber
    Loop While RowNumber <> LastSetRowNumber Or ColumnNumber <> LastSetColumnNumber

BreakLoop:

    If MinPoint < 9999 Then
        ' 抽出した候補を試行
        TryNumberList = BitmapToNumberValues(CurrentStage.MapInfo.Matrix(TryRowNumber, TryColumnNumber).RemainNumberMap)
        
        With TryNumberList
            If CheckInDescendingOrder Then
                StartIndex = .Length
                EndIndex = 1
                StepIndex = -1
            Else
                StartIndex = 1
                EndIndex = .Length
                StepIndex = 1
            End If
            
            For Index = StartIndex To EndIndex Step StepIndex
                Call PushStage(TryRowNumber, TryColumnNumber, .NumberValues(Index))
                
                If TrySudokuMatrix() Then
                    TrySudokuMatrix = True
                    Exit Function
                End If
                
                Call PopStage
            Next
        End With
    End If

Failure:
    Call RestoreStage(SavedStageNumber)
    
    TrySudokuMatrix = False
    Exit Function

CheckLastStage:
    If IsLastStage Then
        TrySudokuMatrix = True
        Exit Function
    End If
    Return
End Function

Private Function GetRemainNumberList(ByRef RowNumber As Long, ByRef ColumnNumber As Long) As NumberValueList
    Dim TargetBitMap As Long
    Dim NumberList As NumberValueList
    Dim GroupNumber As Long
    Dim GroupItemNumber As Long
    Dim CheckNumberList As NumberValueList
    Dim ReservedNumberMap As Long
    Dim CheckRowNumber As Long
    Dim CheckColumnNumber As Long
    Dim CheckGroupNumber As Long
    Dim CheckGroupItemNumber As Long
    
    With CurrentStage.MapInfo
        TargetBitMap = .Matrix(RowNumber, ColumnNumber).RemainNumberMap
        NumberList = BitmapToNumberValues(TargetBitMap)
        
        If NumberList.Length <= 1 Then
            GetRemainNumberList = NumberList
            Exit Function
        End If
        
        ReservedNumberMap = 0
        For CheckRowNumber = 1 To 9
            If CheckRowNumber <> RowNumber Then
                ReservedNumberMap = ReservedNumberMap Or .Matrix(CheckRowNumber, ColumnNumber).RemainNumberMap
            End If
        Next
        GoSub CheckFixed
        
        ReservedNumberMap = 0
        For CheckColumnNumber = 1 To 9
            If CheckColumnNumber <> ColumnNumber Then
                ReservedNumberMap = ReservedNumberMap Or .Matrix(RowNumber, CheckColumnNumber).RemainNumberMap
            End If
        Next
        GoSub CheckFixed
        
        ReservedNumberMap = 0
        Call PositionToGroup(RowNumber, ColumnNumber, GroupNumber, GroupItemNumber)
        For CheckGroupItemNumber = 1 To 9
            If CheckGroupItemNumber <> GroupItemNumber Then
                Call GroupToPosition(GroupNumber, CheckGroupItemNumber, CheckRowNumber, CheckColumnNumber)
                ReservedNumberMap = ReservedNumberMap Or .Matrix(CheckRowNumber, CheckColumnNumber).RemainNumberMap
            End If
        Next
        GoSub CheckFixed
        
        GetRemainNumberList = NumberList
    End With
    
    Exit Function
    
CheckFixed:
    CheckNumberList = BitmapToNumberValues(TargetBitMap And (Not ReservedNumberMap))
    If CheckNumberList.Length = 1 Then
        GetRemainNumberList = CheckNumberList
        Exit Function
    End If
    Return
End Function

Private Sub InitializeStageHistory()
    If StageHistoryInitialized Then Exit Sub
    
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    
    StageHistoryLength = 0
    
    With StageHistory(StageHistoryLength)
        For RowNumber = 1 To 9
            For ColumnNumber = 1 To 9
                With .MapInfo.Matrix(RowNumber, ColumnNumber)
                    .RemainNumberMap = BitAll
                End With
            Next ColumnNumber
        Next RowNumber
    End With
    
    StageLogLength = 0
    ReDim StageLog(1 To 4, 1 To 1)
    
    StageHistoryInitialized = True
End Sub

Private Sub ResetStageHistory()
    StageHistoryLength = 0
    
    StageLogLength = 0
    ReDim StageLog(1 To 4, 1 To 1)
End Sub

Private Sub RestoreStage(ByRef StageNumber As Long)
    Do While StageNumber < StageHistoryLength
        Call PopStage
    Loop
End Sub

Private Sub PushStage(ByRef SetRowNumber As Long, ByRef SetColumnNumber As Long, ByRef SetNumberValue As Long)
    Dim SetGroupNumber As Long
    Dim SetGroupItemNumber As Long
    Dim CheckRowNumber As Long
    Dim CheckColumnNumber As Long
    Dim CheckGroupItemNumber As Long
    
    TryCounter = TryCounter + 1
    
    NumberMatrix(SetRowNumber, SetColumnNumber) = SetNumberValue
    
    StageHistoryLength = StageHistoryLength + 1
    
    If StageLogIsEnabled Then
        StageLogLength = StageLogLength + 1
        ReDim Preserve StageLog(1 To 4, 1 To StageLogLength)
        StageLog(1, StageLogLength) = StageHistoryLength
        StageLog(2, StageLogLength) = SetRowNumber
        StageLog(3, StageLogLength) = SetColumnNumber
        StageLog(4, StageLogLength) = SetNumberValue
    End If
    
    With StageHistory(StageHistoryLength)
        .RowNumber = SetRowNumber
        .ColumnNumber = SetColumnNumber
        .NumberValue = SetNumberValue
        .MapInfo = StageHistory(StageHistoryLength - 1).MapInfo
        
        With .MapInfo.Matrix(SetRowNumber, SetColumnNumber)
            .RemainNumberMap = Bit(SetNumberValue)
        End With
        
        Call PositionToGroup(SetRowNumber, SetColumnNumber, SetGroupNumber, SetGroupItemNumber)
        
        For CheckRowNumber = 1 To 9
            If CheckRowNumber <> SetRowNumber Then
                With .MapInfo.Matrix(CheckRowNumber, SetColumnNumber)
                    Call BitOff(.RemainNumberMap, SetNumberValue)
                End With
            End If
        Next
        
        For CheckColumnNumber = 1 To 9
            If CheckColumnNumber <> SetColumnNumber Then
                With .MapInfo.Matrix(SetRowNumber, CheckColumnNumber)
                    Call BitOff(.RemainNumberMap, SetNumberValue)
                End With
            End If
        Next
        
        For CheckGroupItemNumber = 1 To 9
            If CheckGroupItemNumber <> SetGroupItemNumber Then
                Call GroupToPosition(SetGroupNumber, CheckGroupItemNumber, CheckRowNumber, CheckColumnNumber)
                With .MapInfo.Matrix(CheckRowNumber, CheckColumnNumber)
                    Call BitOff(.RemainNumberMap, SetNumberValue)
                End With
            End If
        Next
    End With
End Sub

Private Sub PopStage()
    With StageHistory(StageHistoryLength)
        NumberMatrix(.RowNumber, .ColumnNumber) = 0
    End With
    
    StageHistoryLength = StageHistoryLength - 1
End Sub

Private Sub InitializeBitOperation(Optional ByRef MaxNumberValue = 9)
    If BitOperationInitialized Then Exit Sub
    
    Dim NumberValue As Long
    Dim BitIndex As Long
    Dim Index As Long
    
    BitAll = 0
    ReDim Bit(1 To MaxNumberValue)
    
    For NumberValue = 1 To MaxNumberValue
        Bit(NumberValue) = 2 ^ (NumberValue - 1)
        Call BitOn(BitAll, NumberValue)
    Next

    ReDim BitmapToNumberValues(0 To BitAll)
    
    For NumberValue = 1 To MaxNumberValue
        BitIndex = Bit(NumberValue)
        
        For Index = 0 To BitIndex - 1
            BitmapToNumberValues(BitIndex + Index) = BitmapToNumberValues(Index)

            With BitmapToNumberValues(BitIndex + Index)
                .Length = .Length + 1
                .NumberValues(.Length) = NumberValue
            End With
        Next Index
    Next NumberValue
    
    BitOperationInitialized = True
End Sub

Private Sub BitOn(ByRef BitMap As Long, ByRef NumberValue As Long)
    BitMap = BitMap Or Bit(NumberValue)
End Sub

Private Sub BitOff(ByRef BitMap As Long, ByRef NumberValue As Long)
    BitMap = BitMap And (Not Bit(NumberValue))
End Sub

Private Sub InitializeGroupMap()
    If GroupMapInitialized Then Exit Sub
    
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim GroupNumber As Long
    Dim GroupItemNumber As Long
    
    For RowNumber = 1 To 9
        For ColumnNumber = 1 To 9
            Call CalcPositionToGroup(RowNumber, ColumnNumber, GroupNumber, GroupItemNumber)
            With PositionToGroupMap(RowNumber, ColumnNumber)
                .GroupNumber = GroupNumber
                .GroupItemNumber = GroupItemNumber
            End With
        Next ColumnNumber
    Next RowNumber
    
    For GroupNumber = 1 To 9
        For GroupItemNumber = 1 To 9
            Call CalcGroupToPosition(GroupNumber, GroupItemNumber, RowNumber, ColumnNumber)
            With GroupToPositionMap(GroupNumber, GroupItemNumber)
                .RowNumber = RowNumber
                .ColumnNumber = ColumnNumber
            End With
        Next GroupItemNumber
    Next GroupNumber
    
    GroupMapInitialized = True
End Sub

Private Sub PositionToGroup(ByRef RowNumber As Long, ByRef ColumnNumber As Long, ByRef GroupNumber As Long, ByRef GroupItemNumber As Long)
    With PositionToGroupMap(RowNumber, ColumnNumber)
        GroupNumber = .GroupNumber
        GroupItemNumber = .GroupItemNumber
    End With
End Sub

Private Sub GroupToPosition(ByRef GroupNumber As Long, ByRef GroupItemNumber As Long, ByRef RowNumber As Long, ByRef ColumnNumber As Long)
    With GroupToPositionMap(GroupNumber, GroupItemNumber)
        RowNumber = .RowNumber
        ColumnNumber = .ColumnNumber
    End With
End Sub

Private Sub CalcPositionToGroup(ByRef RowNumber As Long, ByRef ColumnNumber As Long, ByRef GroupNumber As Long, ByRef GroupItemNumber As Long)
    GroupNumber = 3 * (Fix((RowNumber + 2) / 3) - 1) + Fix((ColumnNumber + 2) / 3)
    GroupItemNumber = 3 * ((RowNumber + 2) Mod 3) + 1 + ((ColumnNumber + 2) Mod 3)
End Sub

Private Sub CalcGroupToPosition(ByRef GroupNumber As Long, ByRef GroupItemNumber As Long, ByRef RowNumber As Long, ByRef ColumnNumber As Long)
    ColumnNumber = 1 + 3 * ((GroupNumber - 1) Mod 3) + ((GroupItemNumber - 1) Mod 3)
    RowNumber = 1 + 3 * Fix((GroupNumber - 1) / 3) + Fix((GroupItemNumber - 1) / 3)
End Sub
'}


