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
