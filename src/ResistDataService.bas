Attribute VB_Name = "ResistDataService"
' RegistDataシートに関するサービスクラス
'

Public Sub initInputData()
' データ入力欄の初期化
'
    Application.ScreenUpdating = False
    
    ' tier欄の初期化
    initInputTier
    ' 形式欄の初期化
    initInputFormat
    ' スタート位置欄の初期化
    initInputStartingRank
    ' コース名欄の初期化
    initInputTrackName
    ' 順位欄の初期化
    initInputRank
    ' 備考欄の初期化
    initInputRemark
    ' コース画像の消去
    removeAllTrackImg
    ' 知識の消去
    initKnowledge
    
    Range(INIT_SELECT_REGIST_DATA).Select
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub initInputTier()
' tier欄の初期化

    ' 初期値の取得
    Dim initValue As String: initValue = Sheets(STORAGE).Cells(1, STORAGE_COL_TIER_NAME).Value
    
    Sheets(REGIST_DATA).Cells(REGIST_ROW_TIER, REGIST_COL_TIER).Value = initValue
    
End Sub

Private Sub initInputFormat()
' 形式欄の初期化

    ' 初期値の取得
    Dim initValue As String: initValue = Sheets(STORAGE).Cells(1, STORAGE_COL_FORMAT_NAME).Value
    
    Sheets(REGIST_DATA).Cells(REGIST_ROW_FORMAT, REGIST_COL_FORMAT).Value = initValue
    
End Sub

Private Sub initInputTrackName()
' コース名欄の初期化
'
    ' 初期値の取得
    Dim initValue As String: initValue = Sheets(STORAGE).Cells(1, STORAGE_COL_TRACK_NAME).Value
    
    Sheets(REGIST_DATA).Select
    
    ' 初期化の実行
    Dim i As Integer
    For i = 1 To RACE_NUM:
        Cells(REGIST_ROW_HEADER + i, REGIST_COL_TRACK_NAME).Value = initValue
    Next i
    
End Sub

Private Sub initInputStartingRank()
' スタート位置欄の初期化
'
    Sheets(REGIST_DATA).Select
    
    Dim i As Integer
    For i = 1 To RACE_NUM
        Cells(REGIST_ROW_HEADER + i, REGIST_COL_START_RANK).Value = ""
    Next i
    
End Sub

Private Sub initInputRank()
' 順位欄の初期化
'
    Sheets(REGIST_DATA).Select
    
    Dim i As Integer
    For i = 1 To RACE_NUM
        Cells(REGIST_ROW_HEADER + i, REGIST_COL_RANK).Value = ""
    Next i
    
End Sub
    
Private Sub initInputRemark()
' 備考欄の初期化
'
    Sheets(REGIST_DATA).Select
    
    Dim i As Integer
    For i = 1 To RACE_NUM
        Cells(REGIST_ROW_HEADER + i, REGIST_COL_REMARK).Value = ""
    Next i
    
End Sub

Public Sub registData()
' データを登録する
'
    ' 入力データの形成
    Dim iData As InputData: Set iData = createInputData
    
    ' データの追加
    Call addNewData(iData)
    
    ' コース画像の消去
    removeAllTrackImg
    
End Sub

Private Function createInputData() As InputData
' 入力データを形成する
'
    ' 登録キー
    Dim registKey As Long: registKey = getNewRegistKey
    ' 日付
    Dim playDate As Date: playDate = Date
    ' tier
    Dim tier As String: tier = Sheets(REGIST_DATA).Cells(REGIST_ROW_TIER, REGIST_COL_TIER).Value
    ' 形式
    Dim format As String: format = Sheets(REGIST_DATA).Cells(REGIST_ROW_FORMAT, REGIST_COL_FORMAT).Value
    ' コースデータ
    Dim track() As TrackData
    Dim arrSize As Integer: arrSize = 0
    
    Dim i As Integer
    For i = 1 To RACE_NUM
        Dim tmpTrack As TrackData: Set tmpTrack = createTrackData(i)
        If Not tmpTrack Is Nothing Then
            ReDim Preserve track(arrSize)
            Set track(arrSize) = tmpTrack
            
            arrSize = arrSize + 1
        End If
    Next i
    
    ' 入力データ
    Dim iData As New InputData
    Call iData.init(registKey, tier, format, Date, track)
    Set createInputData = iData
    
End Function

Private Function createTrackData(i As Integer) As TrackData
' コースデータを形成する
'
    Dim rowNo As Long: rowNo = REGIST_ROW_HEADER + i
    
    ' 入力チェック
    Dim isCompleted As Boolean
    isCompleted = isInputTrackName(i) And isInputRank(i)
    
    If Not isCompleted Then
        Set createTrackData = Nothing
    Else
        Dim track As New TrackData
        Call track.init(getTrackKey(getTrackName(i)), getStartingRank(i), getResultRank(i), getRemark(i))
        Set createTrackData = track
    End If
    
End Function

Private Function isInputTrackName(i As Integer)
' コース名が入力されているか
'
    ' コースを選択の文言
    Dim unselectValue As String: unselectValue = getSelectTrackValue(getLanguage)
    ' 判定対象セル
    Dim c As Range: Set c = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_TRACK_NAME)
    
    isInputTrackName = c.Value <> "" And c.Value <> unselectValue
    
    Debug.Print unselectValue
    
End Function

Private Function isInputRank(i As Integer)
' 順位が入力されているか

    ' 判定対象セル
    Dim c As Range: Set c = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_RANK)
    
    isInputRank = c.Value <> ""
End Function

Private Function getTrackName(i As Integer) As String
' コース名を取得
    getTrackName = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_TRACK_NAME)
End Function

Private Function getResultRank(i As Integer) As Integer
' 結果順位を取得
    getResultRank = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_RANK)
End Function

Private Function getRemark(i As Integer) As String
'備考を取得
    getRemark = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_REMARK)
End Function

Private Function getStartingRank(i As Integer) As Integer
' スタート順位を取得
    getStartingRank = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_START_RANK)
End Function
