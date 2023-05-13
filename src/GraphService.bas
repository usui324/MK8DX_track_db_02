Attribute VB_Name = "GraphService"
Option Explicit

Public Sub updateGraphs()
' グラフを更新する
'
    ActiveWorkbook.RefreshAll
End Sub

Public Sub resetGraphFilter()
' グラフのフィルターをリセットする
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    ' フィルターのリセット
    pTable.PivotFields(PIVOT_FILTER_NAME_1).CurrentPage = "(ALL)"
    pTable.PivotFields(PIVOT_FILTER_NAME_2).CurrentPage = "(ALL)"
    pTable.PivotFields(PIVOT_FILTER_NAME_3).CurrentPage = "(ALL)"
    pTable.PivotFields(PIVOT_FILTER_NAME_4).CurrentPage = "(ALL)"
    
End Sub

Public Sub setGraphMinNumOfRace()
' 規定レース数の設定をかける
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    ' 規定レース数の取得
    Dim reguRaceNum As Long: reguRaceNum = Sheets(SETTINGS).Cells(SETTINGS_ROW_RACE_NUM, SETTINGS_COL_VALUE).Value
    
    ' 設定をかける ' TODO: 行ソースが変わったときの対応
    pTable.PivotFields(PIVOT_ROW_NAME).ClearAllFilters
    pTable.PivotFields(PIVOT_ROW_NAME).PivotFilters. _
        Add2 Type:=xlValueIsGreaterThanOrEqualTo, _
        DataField:=pTable.PivotFields(PIVOT_COL_NAME_3), Value1:=reguRaceNum
End Sub

Public Sub sortGraphByPoint()
' 平均得点でソート
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    ' フラグの値
    Dim flg As Integer: flg = getPointFlg
    
    ' 規定レースフィルター
    setGraphMinNumOfRace
    
    ' フラグの値が0なら昇順ソート / 1なら降順ソート
    If flg = 0 Then
        sortGraphByPointAscending
    Else
        sortGraphByPointDescending
    End If
    
    ' フラグを加算
    incrementPointFlg
    
End Sub

Public Sub sortGraphByRank()
' 平均順位でソート
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    ' 規定レースフィルター
    setGraphMinNumOfRace
    
    ' フラグの値
    Dim flg As Integer: flg = getRankFlg
    
    ' フラグの値が0なら昇順ソート / 1なら降順ソート
    If flg = 0 Then
        sortGraphByRankAscending
    Else
        sortGraphByRankDescending
    End If
    
    ' フラグを加算
    incrementRankFlg
    
End Sub

Public Sub sortGraphByTimes()
' 回数でソート
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    ' フィルターをリセット
    pTable.PivotFields(PIVOT_ROW_NAME).ClearAllFilters
    
    ' フラグの値
    Dim flg As Integer: flg = getTimesFlg
    
    ' フラグの値が0なら昇順ソート / 1なら降順ソート
    If flg = 0 Then
        sortGraphByTimesAscending
    Else
        sortGraphByTimesDescending
    End If
    
    ' フラグを加算
    incrementTimesFlg
    
End Sub

Private Sub sortGraphByPointAscending()
' 平均得点の昇順ソート
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlAscending, PIVOT_COL_NAME_1
    
End Sub

Private Sub sortGraphByPointDescending()
' 平均得点の降順ソート
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlDescending, PIVOT_COL_NAME_1
    
End Sub

Private Sub sortGraphByRankAscending()
' 平均順位の昇順ソート
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlAscending, PIVOT_COL_NAME_2
    
End Sub

Private Sub sortGraphByRankDescending()
' 平均順位の降順ソート
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlDescending, PIVOT_COL_NAME_2
    
End Sub

Private Sub sortGraphByTimesAscending()
' 回数の昇順ソート
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlAscending, PIVOT_COL_NAME_3
    
End Sub

Private Sub sortGraphByTimesDescending()
' 回数の降順ソート
'
    ' ピボットテーブル
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlDescending, PIVOT_COL_NAME_3
    
End Sub
