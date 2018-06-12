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
' オプション用
Private CurrentScreenUpdate As Boolean ' True: 画面更新

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
Private Type BitNumberList
    Length As Long
    NumberValues(1 To 9) As Long
End Type

Private BitOperationInitialized As Boolean
Private Bit() As Long
Private BitAll As Long
Private BitmapToNumberValues() As BitNumberList

' 試行履歴用
Private Type RemainNumberInfo
    NumberValuesBitMap As Long
End Type

Private Type MapInfo
    RemainNumberMap(1 To 9, 1 To 9) As RemainNumberInfo
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

' 数独結果格納用
Private SudokuRange As Range ' 数独用セル範囲
Private NumberMatrix() As Long ' 数独行列格納用
Private TryResult As Boolean ' 試行結果
Private TryCounter As Long ' 試行回数
Private StartTime As Double ' 試行開始タイマ値
Private EndTime As Double ' 試行終了タイマ値
Private ElapsedTime As Double ' 経過時間
Private ElapsedTimeString As String ' 経過時間(Formatted)
'}

'■プロパティ(Public) {
Public Property Let ScreenUpdate(ByVal Flag As Boolean)
    CurrentScreenUpdate = Flag
End Property

Public Property Get ScreenUpdate() As Boolean
    ScreenUpdate = CurrentScreenUpdate
End Property

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
    
    Call InitializeBitOperation
    Call InitializeGroupMap
    Call InitializeStageHistory
End Sub

Private Sub Class_Terminate()
End Sub
'}

'■ メソッド(Public) {
Public Function TrySudoku(ByRef TargetSudokuRange As Range, Optional ByRef Reset As Boolean = False, Optional ByRef Update As Boolean = True) As Boolean
    If Reset Then Call ResetSudokuRange(TargetSudokuRange)
    
    TryCounter = 0
    
    Set SudokuRange = TargetSudokuRange
    
    StartTime = Timer
    
    Call InitializeSudokuMatrix(SudokuRange)
    
    TryResult = TrySudokuMatrix()
    
    EndTime = Timer
    If EndTime < StartTime Then EndTime = EndTime + 24 * 60 * 60
    
    ElapsedTime = EndTime - StartTime
    ElapsedTimeString = Format(ElapsedTime, "0.0000")
    
    If Update Then SudokuRange.Value = NumberMatrix
    
    TrySudoku = TryResult
End Function

Public Function TrySudokuValues(ByRef TargetNumberValues As Variant, ByRef ResultNumberValues As Variant) As Boolean
    TryCounter = 0
    
    StartTime = Timer
    
    Call InitializeSudokuMatrix(TargetNumberValues)
    
    TryResult = TrySudokuMatrix()
    
    EndTime = Timer
    If EndTime < StartTime Then EndTime = EndTime + 24 * 60 * 60
    
    ElapsedTime = EndTime - StartTime
    ElapsedTimeString = Format(ElapsedTime, "0.0000")
    
    ResultNumberValues = NumberMatrix

    TrySudokuValues = TryResult
End Function

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
        
        If 0 <= CheckSudokuMatrix(NumberMatrix, ErrorInfos) Then
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
        
        RemainCounter = CheckSudokuMatrix(NumberMatrix, ErrorInfos)
        
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
Private Function CheckSudokuMatrix(ByRef NumberMatrix() As Long, ByRef ErrorInfos() As Variant) As Long
    Dim RowNumberValueCounters(1 To 9, 1 To 9) As Long
    Dim ColumnNumberValueCounters(1 To 9, 1 To 9) As Long
    Dim GroupNumberValueCounters(1 To 9, 1 To 9) As Long
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim GroupNumber As Long
    Dim GroupItemNumber As Long
    Dim SetNumberValue As Long
    Dim ErrorIndex As Long
        
    CheckSudokuMatrix = 81
    ReDim ErrorInfos(0)
    
    For RowNumber = 1 To 9
        For ColumnNumber = 1 To 9
            SetNumberValue = NumberMatrix(RowNumber, ColumnNumber)
            If SetNumberValue <> 0 Then
                Call PositionToGroup(RowNumber, ColumnNumber, GroupNumber, GroupItemNumber)
                RowNumberValueCounters(RowNumber, SetNumberValue) = RowNumberValueCounters(RowNumber, SetNumberValue) + 1
                ColumnNumberValueCounters(ColumnNumber, SetNumberValue) = ColumnNumberValueCounters(ColumnNumber, SetNumberValue) + 1
                GroupNumberValueCounters(GroupNumber, SetNumberValue) = GroupNumberValueCounters(GroupNumber, SetNumberValue) + 1
            End If
        Next ColumnNumber
    Next RowNumber
    
    For RowNumber = 1 To 9
        For ColumnNumber = 1 To 9
            SetNumberValue = NumberMatrix(RowNumber, ColumnNumber)
            If SetNumberValue <> 0 Then
                Call PositionToGroup(RowNumber, ColumnNumber, GroupNumber, GroupItemNumber)
                
                If _
                    1 < RowNumberValueCounters(RowNumber, SetNumberValue) Or _
                    1 < ColumnNumberValueCounters(ColumnNumber, SetNumberValue) Or _
                    1 < GroupNumberValueCounters(GroupNumber, SetNumberValue) _
                Then
                    ReDim Preserve ErrorInfos(UBound(ErrorInfos) + 1)
                    ErrorInfos(UBound(ErrorInfos)) = Array(RowNumber, ColumnNumber)
                Else
                    CheckSudokuMatrix = CheckSudokuMatrix - 1
                End If
            End If
        Next ColumnNumber
    Next RowNumber
    
    If UBound(ErrorInfos) <= 0 Then
        Exit Function
    End If
    
    CheckSudokuMatrix = -1
End Function

Private Sub InitializeSudokuMatrix(ByRef SudokuRangeValues As Variant)
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim NumberValue As Long
    
    ReDim NumberMatrix(1 To 9, 1 To 9)
    
    Call ResetStageHistory
    
    For RowNumber = 1 To 9
        For ColumnNumber = 1 To 9
            NumberValue = SudokuRangeValues(RowNumber, ColumnNumber)
            If NumberValue <> 0 Then
                Call PushStage(RowNumber, ColumnNumber, NumberValue)
            End If
        Next ColumnNumber
    Next RowNumber
End Sub

Private Function TrySudokuMatrix() As Boolean
    Dim SavedStageNumber As Long
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim CurrentRemainNumberMap() As RemainNumberInfo
    Dim NumberList As BitNumberList
    Dim MinNumberLength As Long
    Dim TryRowNumber As Long
    Dim TryColumnNumber As Long
    Dim TryNumberList As BitNumberList
    Dim Index As Long
    
    If ScreenUpdate Then Range("A1:I9").Value = NumberMatrix
    
    TryCounter = TryCounter + 1
    
    GoSub CheckLastStage
    
    SavedStageNumber = CurrentStageNumber
    
    ' 候補がひとつだけのマスを埋める
    CurrentRemainNumberMap = CurrentStage.MapInfo.RemainNumberMap
    
    For RowNumber = 1 To 9
        For ColumnNumber = 1 To 9
            If NumberMatrix(RowNumber, ColumnNumber) = 0 Then
                NumberList = BitmapToNumberValues(CurrentRemainNumberMap(RowNumber, ColumnNumber).NumberValuesBitMap)
                With NumberList
                    If .Length = 1 Then
                        Call PushStage(RowNumber, ColumnNumber, .NumberValues(1))
                        
                        GoSub CheckLastStage
                        
                        CurrentRemainNumberMap = CurrentStage.MapInfo.RemainNumberMap
                    End If
                End With
            End If
        Next ColumnNumber
    Next RowNumber
    
    ' 候補が最も少ないマスを抽出
    MinNumberLength = 9999
    
    For RowNumber = 1 To 9
        For ColumnNumber = 1 To 9
            If NumberMatrix(RowNumber, ColumnNumber) = 0 Then
                NumberList = BitmapToNumberValues(CurrentRemainNumberMap(RowNumber, ColumnNumber).NumberValuesBitMap)
                With NumberList
                    If .Length < MinNumberLength Then
                        TryRowNumber = RowNumber
                        TryColumnNumber = ColumnNumber
                        MinNumberLength = .Length
                    End If
                End With
            End If
        Next ColumnNumber
    Next RowNumber
    
    If MinNumberLength < 9999 Then
        ' 抽出した候補を試行
        TryNumberList = BitmapToNumberValues(CurrentRemainNumberMap(TryRowNumber, TryColumnNumber).NumberValuesBitMap)
        
        With TryNumberList
            For Index = 1 To .Length
                Call PushStage(TryRowNumber, TryColumnNumber, .NumberValues(Index))
                
                If TrySudokuMatrix() Then
                    TrySudokuMatrix = True
                    Exit Function
                End If
                
                Call PopStage
            Next
        End With
    End If
    
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

Private Sub InitializeStageHistory()
    If StageHistoryInitialized Then Exit Sub
    
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    
    StageHistoryLength = 0
    
    With StageHistory(StageHistoryLength)
        For RowNumber = 1 To 9
            For ColumnNumber = 1 To 9
                With .MapInfo.RemainNumberMap(RowNumber, ColumnNumber)
                    .NumberValuesBitMap = BitAll
                End With
            Next ColumnNumber
        Next RowNumber
    End With
    
    StageHistoryInitialized = True
End Sub

Private Sub ResetStageHistory()
    StageHistoryLength = 0
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
    
    NumberMatrix(SetRowNumber, SetColumnNumber) = SetNumberValue
    
    StageHistoryLength = StageHistoryLength + 1
    
    With StageHistory(StageHistoryLength)
        .RowNumber = SetRowNumber
        .ColumnNumber = SetColumnNumber
        .NumberValue = SetNumberValue
        .MapInfo = StageHistory(StageHistoryLength - 1).MapInfo
        
        With .MapInfo.RemainNumberMap(SetRowNumber, SetColumnNumber)
            .NumberValuesBitMap = 0
        End With
        
        Call PositionToGroup(SetRowNumber, SetColumnNumber, SetGroupNumber, SetGroupItemNumber)
        
        For CheckRowNumber = 1 To 9
            If CheckRowNumber <> SetRowNumber Then
                With .MapInfo.RemainNumberMap(CheckRowNumber, SetColumnNumber)
                    Call BitOff(.NumberValuesBitMap, SetNumberValue)
                End With
            End If
        Next
        
        For CheckColumnNumber = 1 To 9
            If CheckColumnNumber <> SetColumnNumber Then
                With .MapInfo.RemainNumberMap(SetRowNumber, CheckColumnNumber)
                    Call BitOff(.NumberValuesBitMap, SetNumberValue)
                End With
            End If
        Next
        
        For CheckGroupItemNumber = 1 To 9
            If CheckGroupItemNumber <> SetGroupItemNumber Then
                Call GroupToPosition(SetGroupNumber, CheckGroupItemNumber, CheckRowNumber, CheckColumnNumber)
                With .MapInfo.RemainNumberMap(CheckRowNumber, CheckColumnNumber)
                    Call BitOff(.NumberValuesBitMap, SetNumberValue)
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
    GroupNumber = Fix((RowNumber + 2) / 3) + 3 * (Fix((ColumnNumber + 2) / 3) - 1)
    GroupItemNumber = 1 + ((RowNumber + 2) Mod 3) + 3 * ((ColumnNumber + 2) Mod 3)
End Sub

Private Sub CalcGroupToPosition(ByRef GroupNumber As Long, ByRef GroupItemNumber As Long, ByRef RowNumber As Long, ByRef ColumnNumber As Long)
    RowNumber = 1 + 3 * ((GroupNumber - 1) Mod 3) + ((GroupItemNumber - 1) Mod 3)
    ColumnNumber = 1 + 3 * Fix((GroupNumber - 1) / 3) + Fix((GroupItemNumber - 1) / 3)
End Sub
'}
